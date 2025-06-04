const sql = require("mssql");
const { Pinecone } = require("@pinecone-database/pinecone");
const { encoding_for_model, get_encoding } = require('@dqbd/tiktoken');
const OpenAI = require("openai");
const extractTextFromPdf = require("./extractors/pdf");
const extractTextFromDocx = require("./extractors/docx");
const fs = require("fs");
const path = require("path");
const axios = require("axios");
require("dotenv").config();
const { downloadFile } = require("./sharepoint");

const pinecone = new Pinecone({ apiKey: process.env.PINECONE_API_KEY });
const index = pinecone.index(process.env.PINECONE_INDEX_NAME);

const openai = new OpenAI({ apiKey: process.env.OPENAI_API_KEY });

async function deleteFromPinecone(file) {
  const ids = Array.from({ length: file.chunkCount }, (_, i) => `${file.graphid}_chunk_${i}`);
  if (ids && ids.length > 0) {
    if (ids.length > 1) {
      await index.deleteMany(ids);
    } else {
      await index.deleteOne(ids[0]);
    }
  }
}

/**
 * Get token count of a string using OpenAI-compatible tokenization
 * @param {string} text - The text chunk to tokenize
 * @param {string} model - Model name (e.g., 'text-embedding-3-small')
 * @returns {number} - Token count
 */
async function getTokenCount(text, model = 'text-embedding-3-small') {
  // Fallback to base encoding if model-specific encoding isn't available
  const encoding = encoding_for_model(model) ?? get_encoding('cl100k_base');
  const tokens = encoding.encode(text);
  encoding.free(); // Free WASM memory
  return tokens.length;
}

/**
 * Get embeddings using OpenAI embedding model
 * @param {string} text - The text chunk to tokenize
 * @returns {array} - Returns array of floats representing the embedding
 */
async function embedText(text) {
  const res = await openai.embeddings.create({ input: text, model: "text-embedding-3-small" });
  return res.data[0].embedding;
}

/**
 * Turn a large string of text into smaller chunks using the maxWords and overlap parameters
 * @param {string} text 
 * @param {integer} maxWords 
 * @param {integer} overlap 
 * @returns {array} - Returns an array of text chunks
 */
function chunkText(text, maxWords = 400, overlap = 50) {
  const words = text.split(/\s+/);
  const chunks = [];
  const stride = maxWords - overlap;

  for (let i = 0; i < words.length; i += stride) {
    chunks.push(words.slice(i, i + maxWords).join(" "));
  }

  return chunks;
}

/**
 * Recursively chunk text so that no chunk exceeds the model's max token limit.
 * @param {string} text
 * @param {integer} maxWords
 * @param {integer} overlap
 * @param {integer} maxTokens
 * @param {string} model
 * @returns {Promise<array>} - Array of text chunks
 */
async function chunkTextAdaptive(text, maxWords = 400, overlap = 50, maxTokens = 8192, model = 'text-embedding-3-small') {
  const words = text.split(/\s+/);
  const chunks = [];
  const stride = maxWords - overlap;

  for (let i = 0; i < words.length; i += stride) {
    const chunk = words.slice(i, i + maxWords).join(" ");
    const tokenCount = await getTokenCount(chunk, model);
    if (tokenCount > maxTokens && maxWords > 50) {
      // Recursively split this chunk with a smaller maxWords
      const subChunks = await chunkTextAdaptive(chunk, Math.floor(maxWords / 2), Math.floor(overlap / 2), maxTokens, model);
      chunks.push(...subChunks);
    } else {
      chunks.push(chunk);
    }
  }

  return chunks;
}

/** 
 * Process each chunk of text, get embeddings, and upsert to Pinecone
 * @param {object} pool - SQL connection pool
 * @param {array} chunks - Array of text chunks
 * @param {object} file - File object containing metadata
 * @returns {void}
*/
async function processChunks(pool, chunks, file) {
  let chunkLengthGood = true;
  for (let i = 0; i < chunks.length; i++) {
    const tokenCount = await getTokenCount(chunks[i]);
    if (tokenCount > 8192) {
      chunkLengthGood = false;
    }
  }
  if (chunkLengthGood) {  
    for (let i = 0; i < chunks.length; i++) {
      const vector = await embedText(chunks[i]);
      await index.upsert([{
        id: `${file.graphid}_chunk_${i}`,
        values: vector,
        metadata: {
          filename: file.filename,
          filepath: file.filepath,
          fileurl: file.fileurl,
          chunk: chunks[i],
        }
      }]);
    }
    
    await pool.request()
      .input("graphid", sql.NVarChar(200), file.graphid)
      .input("syncStatus", sql.NVarChar(50), "embedded")
      .input("chunkCount", sql.Int, chunks.length)
      .query("UPDATE fileSync SET syncStatus = @syncStatus, chunkCount = @chunkCount, lastEmbeddedDateTime = GETDATE() WHERE graphid = @graphid");
  } else {
    await pool.request()
      .input("graphid", sql.NVarChar(200), file.graphid)
      .input("syncStatus", sql.NVarChar(50), "error_chunkLengthexceeded")
      .input("chunkCount", sql.Int, chunks.length)
      .query("UPDATE fileSync SET syncStatus = @syncStatus, chunkCount = @chunkCount, lastEmbeddedDateTime = GETDATE() WHERE graphid = @graphid");
  }

  console.info(`Processed ${file.graphid}: ${file.filename}`);
};

/**
  * Update the sync status of a file in the database with an error message
  * @param {object} pool - SQL connection pool
  * @param {string} graphid - The unique identifier for the file
  * @param {string} error - The error message to set
  * @returns {void}
*/
async function updateWithError(pool, graphid, error) {
  await pool.request()
    .input("graphid", sql.NVarChar(200), graphid)
    .input("syncStatus", sql.NVarChar(50), error)
    .query("UPDATE fileSync SET syncStatus = @syncStatus WHERE graphid = @graphid");
};

/**
 * Process a single file: download, extract text, chunk, and upsert to Pinecone
 * @param {object} pool - SQL connection pool
 * @param {object} file - File object containing metadata
 * @param {object} client - Microsoft Graph client
 * @returns {void}
*/
async function processFile(pool, file, client) {
  console.info(`Processing ${file.graphid}: ${file.filename}`);
  const tempPath = path.join("tmp", file.filename);
  const res = await downloadFile(client, file.graphid);
  if (res) {
    // fs.writeFileSync(tempPath, res);

    let content = "";
    let pdfError = false;
    let wordError = false;
    if (file.filename.endsWith(".pdf")) {
      try {
        content = await extractTextFromPdf(Buffer.from(res));  
      } catch (pdfErr) {
        pdfError = true;
        console.error(`PDF Error: ${pdfErr}`);
      }
        
    } else if (file.filename.endsWith(".docx")) {
      try {
        content = await extractTextFromDocx(Buffer.from(res));
      } catch (wordErr) {
        wordError = true;
        console.error(`Word Error: ${wordErr}`);
      }
    }
    // fs.unlinkSync(tempPath);
    if (!pdfError && !wordError) {
      const chunks = chunkTextAdaptive(content);
      await processChunks(pool, chunks, file);
    }
    if (pdfError) {
      await updateWithError(pool, file.graphid, "error_pdf");
    }
    if (wordError) {
      await updateWithError(pool, file.graphid, "error_word");
    }
  } else {
    await updateWithError(pool, file.graphid, "error_download");
  }
};

/**
 * Process pending files from the database and upsert to Pinecone
 * @param {object} pool - SQL connection pool
 * @param {object} client - Microsoft Graph client
 * @returns {void}
*/
async function processPendingFiles(pool, client) {
  const results = await pool.request()
    .query("SELECT * FROM fileSync WHERE agentId = 1 AND isDeleted = 0 AND ignoreFile = 0 AND (syncStatus = 'pending' OR syncStatus = 'error_download' OR syncStatus = 'error_chunkLengthexceeded')");

  for (const file of results.recordset) {
    await processFile(pool, file, client);
  }
}

/**
 * Reprocess updated files from the database and upsert to Pinecone
 * @param {object} pool - SQL connection pool
 * @param {object} client - Microsoft Graph client
 * @returns {void}
*/
async function reprocessUpdatedFiles(pool, client) {
  const results = await pool.request().query(`
    SELECT * FROM fileSync WHERE agentId = 1 AND isDeleted = 0 AND ignoreFile = 0 AND syncStatus = 'updated'
  `);

  for (const file of results.recordset) {
    await deleteFromPinecone(file);
    await processFile(pool, file, client);
  }
}

/**
 * Remove deleted files from Pinecone and update the database
 * @param {object} pool - SQL connection pool
 * @returns {void}
*/
async function removeDeletedFiles(pool) {
  const results = await pool.request()
    .query("SELECT graphid, chunkCount FROM fileSync WHERE agentId = 1 AND (isDeleted = 1 or ignoreFile = 1) AND chunkCount IS NOT NULL AND chunkCount > 0");

  for (const row of results.recordset) {
    console.info('Row:', row);
    await deleteFromPinecone(row);
    await pool.request()
      .input("graphid", sql.NVarChar(200), row.graphid)
      .input("syncStatus", sql.NVarChar(50), "pending")
      .input("chunkCount", sql.Int, 0)
      .query("UPDATE fileSync SET syncStatus = @syncStatus, chunkCount = @chunkCount, lastEmbeddedDateTime = GETDATE() WHERE graphid = @graphid");

    console.info(`Removed deleted file: ${row.graphid}`);
  }
}

module.exports = { processPendingFiles, reprocessUpdatedFiles, removeDeletedFiles };
