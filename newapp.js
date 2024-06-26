app.js
  

const express = require('express');
const fs = require('fs');
const path = require('path');
const { processDocument } = require('./postdata');
const app = express();

app.use(express.json());

const log = (message) => {
  console.log(`[${new Date().toISOString()}] ${message}`);
};

app.post('/update-document', async (req, res) => {
  try {
    log(`Received JSON: ${JSON.stringify(req.body)}`);
    const jsonContent = req.body;

    // Generate the updated document
    const updatedDocumentPath = await processDocument(jsonContent);

    // Send the file as a response
    res.download(updatedDocumentPath, 'updated_document.docx', (err) => {
      if (err) {
        log(`Error sending file: ${err.message}`);
        res.status(500).send('Error sending file');
      }

      // Clean up the file after sending
      fs.unlink(updatedDocumentPath, (err) => {
        if (err) log(`Error deleting file: ${err.message}`);
      });
    });
  } catch (error) {
    log(`Error processing document: ${error.message}`);
    res.status(500).send('Error processing document');
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  log(`Server is running on port ${PORT}`);
});



postdata.js


const fs = require('fs');
const axios = require('axios');
const AdmZip = require('adm-zip');
const xpath = require('xpath');
const { DOMParser, XMLSerializer } = require('@xmldom/xmldom');
const { v4: uuidv4 } = require('uuid');

const log = (message) => {
  console.log(`[${new Date().toISOString()}] ${message}`);
};

const readZipFile = (zipFilePath, fileName) => {
  const zip = new AdmZip(zipFilePath);
  const zipEntry = zip.getEntry(fileName);
  const content = zipEntry.getData().toString('utf8');
  return content;
};

const writeZipFile = (zipFilePath, fileName, content) => {
  const zip = new AdmZip(zipFilePath);
  zip.updateFile(fileName, Buffer.from(content, 'utf8'));
  zip.writeZip(zipFilePath);
};

const replacePlaceholder = (documentContent, jsonContent) => {
  log(`Entering replacePlaceholder`);
  const doc = new DOMParser().parseFromString(documentContent, 'text/xml');
  const select = xpath.useNamespaces({ "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main" });

  select("//w:sdt", doc).forEach(node => {
    const tagNode = select('.//w:tag/@w:val', node)[0];
    if (tagNode) {
      const tagName = tagNode.value;
      log(`Processing tag: ${tagName}`);
      const tagValue = jsonContent[tagName];

      if (tagValue !== null && tagValue !== undefined) {
        log(`Found value for ${tagName}: ${tagValue}`);
        const textNodes = select('.//w:t', node);

        if (textNodes.length > 0) {
          log(`Replacing text nodes for tag: ${tagName}`);
          textNodes.forEach(textNode => {
            textNode.textContent = tagValue;
          });
        } else {
          log(`No text nodes found for tag: ${tagName}`);
        }
      } else {
        log(`No value found for tag: ${tagName}`);
      }
    } else {
      log(`No tagNode found in node`);
    }
  });

  const serializer = new XMLSerializer();
  const updatedContent = serializer.serializeToString(doc);
  log(`Exiting replacePlaceholder`);
  return updatedContent;
};

const processDocument = async (jsonContent) => {
  const templateFilePath = './Test Document.docx';
  const fileName = 'word/document.xml';

  const newFileName = `./updated_${uuidv4()}.docx`;
  fs.copyFileSync(templateFilePath, newFileName);

  const documentContent = readZipFile(newFileName, fileName);
  const updatedContent = replacePlaceholder(documentContent, jsonContent);
  writeZipFile(newFileName, fileName, updatedContent);

  log(`Document updated successfully. Saved as ${newFileName}`);
  return newFileName;
};

module.exports = {
  processDocument,
};
