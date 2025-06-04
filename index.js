const sql = require("mssql");
const { Client } = require("@microsoft/microsoft-graph-client");
require("dotenv").config();
const { getAccessToken, syncDelta } = require("./sharepoint");
const { processPendingFiles, reprocessUpdatedFiles, removeDeletedFiles } = require("./pinecone");

const authProvider = {
  getAccessToken: async () => {
    const accessToken = await getAccessToken();
    return accessToken;
  }
};

(async () => {
  const client = Client.initWithMiddleware({ authProvider });

  const pool = await sql.connect({
    user: process.env.DB_USER,
    password: process.env.DB_PASS,
    database: process.env.DB_DATABASE,
    server: process.env.DB_HOST,
    options: {
      encrypt: true,
      trustServerCertificate: true
    }
  });
  console.info('client:', client);
  // Sync delta changes from SharePoint
  console.info('Syncing delta changes from SharePoint...');
  await syncDelta(client, pool);
  
  // Process Pending Files to Pinecone
  console.info('Processing pending files to Pinecone...');
  await processPendingFiles(pool, client);

  // Process Updated Files to Pinecone
  console.info('Processing updated files to Pinecone...');
  await reprocessUpdatedFiles(pool);

  // Remove Deleted Files from Pinecone
  console.info('Removing deleted files from Pinecone...');
  await removeDeletedFiles(pool);

  console.log("✅ Delta sync complete.");
})();