const { ClientSecretCredential } = require("@azure/identity");
const { ResponseType } = require("@microsoft/microsoft-graph-client");
const sql = require("mssql");
const fs = require("fs");
const axios = require("axios");
require("isomorphic-fetch");
require("dotenv").config();

const deltaStatePath = "./deltaLink.txt";

let count = 1;

function updateCounter() {
  process.stdout.clearLine(0);  // Clear current line
  process.stdout.cursorTo(0);   // Move cursor to beginning of line
  process.stdout.write(`Processing Item #: ${count}`); // Write updated counter
  count++;
}

const credential = new ClientSecretCredential(
  process.env.TENANT_ID,
  process.env.CLIENT_ID,
  process.env.CLIENT_SECRET
);

async function getAccessToken() {
  const token = await credential.getToken("https://graph.microsoft.com/.default");
  return token.token;
}

async function downloadFile(client, graphid) {
  try {
    const fullUrl = `//drives/${process.env.DRIVE_ID}/items/${graphid}/content`;
    const resp = await client
	    .api(fullUrl)
      .responseType(ResponseType.ARRAYBUFFER)
      .get();
    return Buffer.from(resp);
  } catch (error) {
    console.error("Error downloading file:", error);
    return false;
  }
}

// async function downloadFile(client, url) {
//   try {
//     const path = encodeURI(url);
//     const fullUrl = `/sites/${process.env.SHAREPOINT_SITE_ID}/drives/${process.env.DRIVE_ID}/root:/${path}:/content`;
//     const resp = await client
// 	    .api(fullUrl)
//       .responseType(ResponseType.ARRAYBUFFER)
//       .get();
//     return Buffer.from(resp);
//   } catch (error) {
//     console.error("Error downloading file:", error);
//     return false;
//   }
// }

// async function downloadFile(url, localPath) {
//   const res = await axios.get(url, { responseType: "arraybuffer" });
//   fs.writeFileSync(localPath, res.data);
// }

async function syncDelta(client, pool) {
  let deltaUrl = fs.existsSync(deltaStatePath)
    ? fs.readFileSync(deltaStatePath, "utf8")
    : `/drives/${process.env.DRIVE_ID}/root/delta?$select=id,name,size,webUrl,createdDateTime,lastModifiedDateTime,createdBy,lastModifiedBy,file,parentReference`;

  let hasMore = true;

  while (hasMore) {
    const res = await client.api(deltaUrl).get();
    const changes = res.value;

    for (const item of changes) {
      //console.info('Processing Item:', item);
      const isFile = item.file !== undefined;
      const isValid = (isFile && item.name !== undefined);

      if (item.deleted) {
        await pool.request()
          .input("graphid", sql.NVarChar(200), item.id)
          .query("UPDATE fileSync SET isDeleted = 1 WHERE graphid = @graphid");
        continue;
      }

      if (isValid && (item.name.endsWith(".pdf") || item.name.endsWith(".docx"))) {
        updateCounter();
        const exists = await pool.request()
          .input("graphid", sql.NVarChar(200), item.id)
          .query("SELECT id FROM fileSync WHERE graphid = @graphid");

        const path = item.parentReference?.path?.replace(/^.*?:/, "") + "/" + item.name;

        if (exists.recordset.length > 0) {
          await pool.request()
            .input("graphid", sql.NVarChar(200), item.id)
            .input("filename", sql.NVarChar(255), item.name)
            .input("filepath", sql.NVarChar(sql.MAX), path)
            .input("fileurl", sql.NVarChar(sql.MAX), item.webUrl)
            .input("filesize", sql.Int, item.size || 0)
            .input("createdDateTime", sql.DateTime, item.createdDateTime || null)
            .input("createdBy", sql.NVarChar(255), item.createdBy?.user?.displayName || null)
            .input("lastModifiedDateTime", sql.DateTime, item.lastModifiedDateTime || null)
            .input("lastModifiedBy", sql.NVarChar(255), item.lastModifiedBy?.user?.displayName || null)
            .query(`
              UPDATE fileSync SET
                filename = @filename,
                filepath = @filepath,
                fileurl = @fileurl,
                filesize = @filesize,
                createdDateTime = @createdDateTime,
                createdBy = @createdBy,
                lastModifiedDateTime = @lastModifiedDateTime,
                lastModifiedBy = @lastModifiedBy,
                isDeleted = 0,
                syncStatus = 'updated'
              WHERE graphid = @graphid
            `);
        } else {
          await pool.request()
            .input("agentId", sql.Int, 1)
            .input("graphid", sql.NVarChar(200), item.id)
            .input("filename", sql.NVarChar(255), item.name)
            .input("filepath", sql.NVarChar(sql.MAX), path)
            .input("fileurl", sql.NVarChar(sql.MAX), item.webUrl)
            .input("filesize", sql.Int, item.size || 0)
            .input("createdDateTime", sql.DateTime, item.createdDateTime || null)
            .input("createdBy", sql.NVarChar(255), item.createdBy?.user?.displayName || null)
            .input("lastModifiedDateTime", sql.DateTime, item.lastModifiedDateTime || null)
            .input("lastModifiedBy", sql.NVarChar(255), item.lastModifiedBy?.user?.displayName || null)
            .input("isDeleted", sql.Bit, 0)
            .input("syncStatus", sql.NVarChar(50), "pending")
            .query(`
              INSERT INTO fileSync (
                agentId, graphid, filename, filepath, fileurl, filesize,
                createdDateTime, createdBy,
                lastModifiedDateTime, lastModifiedBy,
                isDeleted, syncStatus
              ) VALUES (
                @agentId, @graphid, @filename, @filepath, @fileurl, @filesize,
                @createdDateTime, @createdBy,
                @lastModifiedDateTime, @lastModifiedBy,
                @isDeleted, @syncStatus
              )
            `);
        }
      }
    }

    if (res["@odata.nextLink"]) {
      deltaUrl = res["@odata.nextLink"];
    } else {
      hasMore = false;
      if (res["@odata.deltaLink"]) {
        fs.writeFileSync(deltaStatePath, res["@odata.deltaLink"]);
      }
    }
  }
  console.info('\nDelta sync complete.');
}

module.exports = { getAccessToken, downloadFile, syncDelta };