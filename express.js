const express = require('express');
const fs = require('fs');
const AdmZip = require('adm-zip');
const xpath = require('xpath');
const { DOMParser, XMLSerializer } = require('@xmldom/xmldom');

const app = express();
app.use(express.json());

const log = (message) => {
  console.log(`[${new Date().toISOString()}] ${message}`);
};

const readZipFile = (zipFilePath, fileName) => {
  log(`Entering readZipFile with zipFilePath=${zipFilePath} and fileName=${fileName}`);
  const zip = new AdmZip(zipFilePath);
  const zipEntry = zip.getEntry(fileName);
  const content = zipEntry.getData().toString('utf8');
  log(`Exiting readZipFile with content=${content.substring(0, 100)}`);
  return content;
};

const writeZipFile = (zipFilePath, fileName, content) => {
  log(`Entering writeZipFile with zipFilePath=${zipFilePath}, fileName=${fileName}, and content=${content.substring(0, 100)}`);
  const zip = new AdmZip(zipFilePath);
  zip.updateFile(fileName, Buffer.from(content, 'utf8'));
  zip.writeZip(zipFilePath);
  log(`Exiting writeZipFile`);
};

const replacePlaceholder = (documentContent, jsonContent) => {
  log(`Entering replacePlaceholder with documentContent=${documentContent.substring(0, 100)} and jsonContent=${JSON.stringify(jsonContent)}`);
  const doc = new DOMParser().parseFromString(documentContent, 'text/xml');
  const select = xpath.useNamespaces({ "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main" });
  const tags = jsonContent;

  select("//w:sdt", doc).forEach(node => {
    var tagNode = select('.//w:tag/@w:val', node)[0];
    if (tagNode) {
      log(`Found tagNode with value=${tagNode.value}`);
      const tagName = tagNode.value;
      if (tags.hasOwnProperty(tagName)) {
        log(`Found matching tag in jsonContent: ${tagName}=${tags[tagName]}`);
        const textNodes = select('.//w:t', node);
        if (textNodes.length > 0) {
          log(`Found text nodes, replacing with ${tags[tagName]}`);
          textNodes.forEach(textNode => {
            textNode.textContent = tags[tagName];
          });
        } else {
          log(`No text nodes found`);
        }
      } else {
        log(`No matching tag in jsonContent: ${tagName}`);
      }
    } else {
      log(`No tagNode found`);
    }
  });
  const serializer = new XMLSerializer();
  const updatedContent = serializer.serializeToString(doc);
  log(`Exiting replacePlaceholder with updatedContent=${updatedContent.substring(0, 100)}`);
  return updatedContent;
};

app.post('/update-document', (req, res) => {
  log(`Entering /update-document with req.body=${JSON.stringify(req.body)}`);
  const zipFilePath = './Test Document.docx';
  const fileName = 'word/document.xml';
  const jsonContent = req.body;

  try {
    log(`Reading zip file`);
    const documentContent = readZipFile(zipFilePath, fileName);
    log(`Replacing placeholders`);
    const updatedContent = replacePlaceholder(documentContent, jsonContent);
    log(`Writing updated content to zip file`);
    writeZipFile(zipFilePath, fileName, updatedContent);

    res.status(200).send('Document updated successfully');
    log(`Exiting /update-document with success`);
  } catch (error) {
    console.error('Error updating document:', error);
    res.status(500).send('Internal Server Error');
    log(`Exiting /update-document with error`);
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  log(`Server is running on port ${PORT}`);
});
