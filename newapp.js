app.js
  
const express = require('express');
const { processDocument } = require('./postdata');
const app = express();

app.use(express.json());

const log = (message) => {
  console.log(`[${new Date().toISOString()}] ${message}`);
};

app.post('/update-document', async (req, res) => {
  log(`Received JSON: ${JSON.stringify(req.body)}`);
  const jsonContent = req.body;

  try {
    const updatedFilePath = await processDocument(jsonContent);
    res.download(updatedFilePath, (err) => {
      if (err) {
        console.error('Error sending file:', err);
        res.status(500).send('Error sending file');
      } else {
        log(`Document updated successfully. Sent as ${updatedFilePath}`);
        // Clean up the file after sending it
        fs.unlink(updatedFilePath, (unlinkErr) => {
          if (unlinkErr) {
            console.error('Error deleting file:', unlinkErr);
          }
        });
      }
    });
  } catch (error) {
    console.error('Error processing document:', error);
    res.status(500).send('Error processing document');
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  log(`Server is running on port ${PORT}`);
});

postdata.js


const fs = require('fs');
const AdmZip = require('adm-zip');
const { v4: uuidv4 } = require('uuid');
const { DOMParser, XMLSerializer } = require('@xmldom/xmldom');
const xpath = require('xpath');

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
  return