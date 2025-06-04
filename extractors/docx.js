const fs = require("fs");
const mammoth = require("mammoth");

// module.exports = async function extractTextFromDocx(filePath) {
//   const buffer = fs.readFileSync(filePath);
//   const result = await mammoth.extractRawText({ buffer });
//   return result.value;
// };

module.exports = async function extractTextFromDocx(dataBuffer) {
  //const buffer = fs.readFileSync(filePath);
  const result = await mammoth.extractRawText({ buffer: dataBuffer });
  dataBuffer.fill(0); // Free memory
  return result.value;
};