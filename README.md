# TGW GPT Embed

This project synchronizes files from Microsoft SharePoint, extracts their text content, chunks and embeds the text using OpenAI, and upserts the embeddings into Pinecone for semantic search. It also manages file sync status in a SQL Server database.

## Features
- Syncs PDF and DOCX files from SharePoint using Microsoft Graph API
- Extracts text from files using `pdf-parse` and `mammoth`
- Chunks text adaptively to fit OpenAI embedding model token limits
- Generates embeddings with OpenAI's `text-embedding-3-small` model
- Upserts embeddings into Pinecone vector database
- Tracks file sync and processing status in SQL Server
- Handles file updates and deletions

## Project Structure
- `index.js` — Main entry point; orchestrates sync, processing, and cleanup
- `pinecone.js` — Handles chunking, embedding, Pinecone upserts, and sync status updates
- `sharepoint.js` — Handles SharePoint delta sync and file downloads
- `extractors/` — Contains file extractors for PDF (`pdf.js`) and DOCX (`docx.js`)
- `.env` — Environment variables (credentials, API keys, DB connection, etc.)

## Environment Variables
Create a `.env` file in the project root with the following variables:

```
TENANT_ID=...           # Azure AD tenant ID
CLIENT_ID=...           # Azure AD app client ID
CLIENT_SECRET=...       # Azure AD app client secret
SHAREPOINT_SITE_ID=...  # SharePoint site ID
DRIVE_ID=...            # SharePoint drive ID
OPENAI_API_KEY=...      # OpenAI API key
PINECONE_API_KEY=...    # Pinecone API key
PINECONE_INDEX_NAME=... # Pinecone index name
DB_USER=...             # SQL Server username
DB_PASS=...             # SQL Server password
DB_HOST=...             # SQL Server host
DB_DATABASE=...         # SQL Server database name
```

**Note:** The `.env` file is excluded from version control via `.gitignore`.

## Usage
1. Install dependencies:
   ```sh
   npm install
   ```
2. Set up your `.env` file as described above.
3. Run the main script:
   ```sh
   node index.js
   ```

## Dependencies
- `@microsoft/microsoft-graph-client` — Microsoft Graph API client
- `@azure/identity` — Azure authentication
- `mssql` — SQL Server client
- `pdf-parse` — PDF text extraction
- `mammoth` — DOCX text extraction
- `openai` — OpenAI API client
- `@pinecone-database/pinecone` — Pinecone vector DB client
- `dotenv`, `axios`, `uuid`, etc.

## License
ISC

## Author
Ryan Schultz <ryan_schultz@graycor.com>
