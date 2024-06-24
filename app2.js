const express = require('express');
const fs = require('fs');
const AdmZip = require('adm-zip');
const xpath = require('xpath');
const { DOMParser } = require('@xmldom/xmldom');

const app = express();
app.use(express.json());

const log = (message) => {
  console.log(`[${new Date().toISOString()}] ${message}`);
};

const readZipFile = (zipFilePath, fileName) => {
  const zip = new AdmZip(zipFilePath);
  const zipEntry = zip.getEntry(fileName);
  if (!zipEntry) {
    throw new Error(`File ${fileName} not found in ${zipFilePath}`);
  }
  return zipEntry.getData().toString('utf8');
};

const extractJSONFromDocument = (zipFilePath, fileName) => {
  try {
    const documentContent = readZipFile(zipFilePath, fileName);
    const doc = new DOMParser().parseFromString(documentContent);
    const select = xpath.useNamespaces({ 'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main' });

    // Example: Extracting text inside <w:t> tags
    const textNodes = select('//w:t/text()', doc);
    const textContent = textNodes.map(node => node.data).join('');

    return { extractedData: textContent };  // Modify as per your document structure
  } catch (error) {
    console.error('Error reading or parsing the document:', error);
    throw error;  // Propagate error to handle in the route handler
  }
};

// Endpoint to extract JSON from Word document
app.get('/extract-json', (req, res) => {
  const zipFilePath = './Test Document.docx';
  const fileName = 'word/document.xml';

  try {
    const extractedData = extractJSONFromDocument(zipFilePath, fileName);
    res.json(extractedData);
  } catch (error) {
    console.error('Error extracting JSON from document:', error);
    res.status(500).send('Error extracting JSON from document');
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  log(`Server is running on port ${PORT}`);
});
