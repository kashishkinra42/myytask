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
======================================

=================================


    =================

const fs = require('fs');
const AdmZip = require('adm-zip');
const xpath = require('xpath');
const dom = require('@xmldom/xmldom').DOMParser;
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
};

const createTextNode = (doc, text, highlighted = false) => {
    const runNode = doc.createElement('w:r');
    if (highlighted) {
        const rPrNode = doc.createElement('w:rPr');
        const highlightNode = doc.createElement('w:highlight');
        highlightNode.setAttribute('w:val', 'yellow');
        rPrNode.appendChild(highlightNode);
        runNode.appendChild(rPrNode);
    }
    const textNode = doc.createElement('w:t');
    textNode.textContent = text;
    runNode.appendChild(textNode);
    return runNode;
};

const createListNode = (doc, items) => {
    return items.map((item, index) => {
        const pNode = doc.createElement('w:p');
        
        const pPrNode = doc.createElement('w:pPr');
        const numPrNode = doc.createElement('w:numPr');
        const numIdNode = doc.createElement('w:numId');
        numIdNode.setAttribute('w:val', '1');
        const ilvlNode = doc.createElement('w:ilvl');
        ilvlNode.setAttribute('w:val', '0');
        numPrNode.appendChild(ilvlNode);
        numPrNode.appendChild(numIdNode);
        pPrNode.appendChild(numPrNode);
        pNode.appendChild(pPrNode);

        if (typeof item === 'string') {
            pNode.appendChild(createTextNode(doc, `${index + 1}. ${item}`));
        } else if (typeof item === 'object' && item.text) {
            pNode.appendChild(createTextNode(doc, `${index + 1}. ${item.text}`, item.highlighted));
        }
        return pNode;
    });
};

const replacePlaceholder = (documentContent, jsonContent) => {
    const doc = new dom().parseFromString(documentContent, 'text/xml');
    const select = xpath.useNamespaces({ "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main" });
    const tags = JSON.parse(jsonContent);

    select("//w:sdt", doc).forEach(node => {
        const tagNode = select('.//w:tag/@w:val', node)[0];
        if (tagNode) {
            const tagName = tagNode.value;
            if (tags.hasOwnProperty(tagName)) {
                const contentNode = select('.//w:sdtContent', node)[0];
                while (contentNode.firstChild) {
                    contentNode.removeChild(contentNode.firstChild); // Clear existing content
                }

                const tagValue = tags[tagName];

                if (Array.isArray(tagValue)) {
                    const listNodes = createListNode(doc, tagValue);
                    listNodes.forEach(listNode => contentNode.appendChild(listNode));
                } else if (typeof tagValue === 'object' && tagValue.text) {
                    const textNode = createTextNode(doc, tagValue.text, tagValue.highlighted);
                    contentNode.appendChild(textNode);
                } else if (typeof tagValue === 'string') {
                    const textNode = createTextNode(doc, tagValue);
                    contentNode.appendChild(textNode);
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

==================
    ===================
    ======================
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

const createNumberedList = (doc, items) => {
    const listXml = items.map((item, index) => {
        return `
            <w:p>
                <w:pPr>
                    <w:numPr>
                        <w:ilvl w:val="0"/>
                        <w:numId w:val="1"/>
                    </w:numPr>
                </w:pPr>
                <w:r>
                    <w:t>${item.highlighted ? `[${item.text}]` : item}</w:t>
                </w:r>
            </w:p>`;
    }).join('');
    return new dom().parseFromString(listXml, 'text/xml').documentElement;
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
                    if (Array.isArray(tags[tagName])) {
                        const listNode = createNumberedList(doc, tags[tagName]);
                        node.parentNode.replaceChild(listNode, node);
                    } else {
                        textNodes.forEach(textNode => {
                            textNode.textContent = tags[tagName];
                        });
                    }
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
