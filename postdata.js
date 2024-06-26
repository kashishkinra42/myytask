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

const processTagValue = (value) => {
  if (Array.isArray(value)) {
    return value.map((item, index) => {
      if (typeof item === 'string') {
        return `${index + 1}. ${item}`;
      } else if (typeof item === 'object' && item.text) {
        return `${index + 1}. ${item.text}`;
      }
      return '';
    }).join('\n');
  } else if (typeof value === 'object' && value.text) {
    return value.text;
  }
  return value;
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
      if (jsonContent.hasOwnProperty(tagName)) {
        const textNodes = select('.//w:t', node);
        if (textNodes.length > 0) {
          const processedValue = processTagValue(jsonContent[tagName]);
          const parts = processedValue.split('\n');

          // Clear existing text nodes
          textNodes.forEach((textNode, index) => {
            if (index === 0) {
              textNode.textContent = parts[0];
            } else {
              textNode.parentNode.removeChild(textNode);
            }
          });

          // Add new text nodes and line breaks
          for (let i = 1; i < parts.length; i++) {
            const breakNode = doc.createElement('w:br');
            const newTextNode = doc.createElement('w:t');
            newTextNode.textContent = parts[i];
            textNodes[0].parentNode.appendChild(breakNode);
            textNodes[0].parentNode.appendChild(newTextNode);
          }
        }
      }
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

  return newFileName;
};

module.exports = { processDocument };
