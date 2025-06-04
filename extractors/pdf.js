const fs = require("fs");
const pdfParse = require("pdf-parse");

// module.exports = async function extractTextFromPdf(filePath) {
//   const dataBuffer = fs.readFileSync(filePath);
//   const data = await pdfParse(dataBuffer);
//   return data.text;
// };

module.exports = async function extractTextFromPdf(dataBuffer) {
  // const dataBuffer = fs.readFileSync(filePath);
  const data = await pdfParse(dataBuffer);
  dataBuffer.fill(0); // Free memory
  return data.text;
};