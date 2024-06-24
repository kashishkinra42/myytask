const fs = require('fs');
const AdmZip = require('adm-zip');
var xpath = require('xpath');
const { json } = require('express/lib/response');
const zipFile = require('adm-zip/zipFile');
var dom = require('@xmldom/xmldom').DOMParser;
const xmlSerializer = require('@xmldom/xmldom').XMLSerializer;

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
}

const replacePlaceholder = (documentContent, jsonContent) => {
    const doc = new dom().parseFromString(documentContent, 'text/xml');
    const select = xpath.useNamespaces({ "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main" });
    const tags = JSON.parse(jsonContent);

    select("//w:sdt", doc).forEach(node => {

        var tagNode = select('.//w:tag/@w:val', node)[0];

        if(tagNode){
            const tagName = tagNode.value;
            if(tags.hasOwnProperty(tagName)){
                const textNodes = select('.//w:t', node);
                if(textNodes.length >0){
                    textNodes.forEach(textNode => {
                        textNode.textContent = tags[tagName];
                    });
                }
            }
        }
    });
    const serializer = new xmlSerializer();
    return serializer.serializeToString(doc);
};


const zipFilePath = './Test Document.docx';
const fileName = 'word/document.xml';
const jsonPath = './myoutput.json';
const content = readZipFile(zipFilePath, fileName);

const documentContent = readZipFile(zipFilePath, fileName);
const jsonContent = fs.readFileSync(jsonPath, 'utf8');

const updatedContent = replacePlaceholder(documentContent, jsonContent);
writeZipFile(zipFilePath, fileName, updatedContent);

console.log('Document updated successfully');



//-------
const fs = require('fs');
const AdmZip = require('adm-zip');
var xpath = require('xpath');
const { json } = require('express/lib/response');
var dom = require('@xmldom/xmldom').DOMParser;
const xmlSerializer = require('@xmldom/xmldom').XMLSerializer;

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
}

const replacePlaceholder = (documentContent, jsonContent) => {
    const doc = new dom().parseFromString(documentContent, 'text/xml');
    const select = xpath.useNamespaces({ "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main" });
    const tags = JSON.parse(jsonContent);

    select("//w:sdt", doc).forEach(node => {
        var tagNode = select('.//w:tag/@w:val', node)[0];

        if (tagNode) {
            const tagName = tagNode.value;
            if (tags.hasOwnProperty(tagName)) {
                const textNodes = select('.//w:t', node);
                if (textNodes.length > 0) {
                    textNodes.forEach(textNode => {
                        const parentNode = textNode.parentNode;
                        const textParts = tags[tagName].split('\n');
                        
                        // Remove original text node
                        parentNode.removeChild(textNode);
                        
                        // Add new text nodes with breaks
                        textParts.forEach((part, index) => {
                            if (index > 0) {
                                // Add a break line element
                                const breakNode = doc.createElementNS("http://schemas.openxmlformats.org/wordprocessingml/2006/main", 'w:br');
                                parentNode.appendChild(breakNode);
                            }
                            // Add text part
                            const newText = doc.createElementNS("http://schemas.openxmlformats.org/wordprocessingml/2006/main", 'w:t');
                            newText.textContent = part;
                            parentNode.appendChild(newText);
                        });
                    });
                }
            }
        }
    });
    const serializer = new xmlSerializer();
    return serializer.serializeToString(doc);
};

const zipFilePath = './Test Document.docx';
const fileName = 'word/document.xml';
const jsonPath = './myoutput.json';
const content = readZipFile(zipFilePath, fileName);

const documentContent = readZipFile(zipFilePath, fileName);
const jsonContent = fs.readFileSync(jsonPath, 'utf8');

const updatedContent = replacePlaceholder(documentContent, jsonContent);
writeZipFile(zipFilePath, fileName, updatedContent);

console.log('Document updated successfully');




