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

const replacePlaceholder = (documentContent, jsonContent) => {
    const doc = new dom().parseFromString(documentContent, 'text/xml');
    const select = xpath.useNamespaces({ "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main" });
    const tags = JSON.parse(jsonContent);

    const processTagValue = (value) => {
        if (Array.isArray(value)) {
            return value.map(item => {
                if (typeof item === 'string') {
                    return item;
                } else if (typeof item === 'object' && item.text) {
                    return item.highlighted ? `**${item.text}**` : item.text;
                }
                return '';
            }).join(', ');
        } else if (typeof value === 'object' && value.text) {
            return value.highlighted ? `**${value.text}**` : value.text;
        }
        return value;
    };

    select("//w:sdt", doc).forEach(node => {
        var tagNode = select('.//w:tag/@w:val', node)[0];

        if(tagNode){
            const tagName = tagNode.value;
            if(tags.hasOwnProperty(tagName)){
                const textNodes = select('.//w:t', node);
                if(textNodes.length >0){
                    const processedValue = processTagValue(tags[tagName]);
                    textNodes.forEach(textNode => {
                        textNode.textContent = processedValue;
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

========================
====================

    const fs = require('fs');
const AdmZip = require('adm-zip');
var xpath = require('xpath');
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
};

const replacePlaceholder = (documentContent, jsonContent) => {
    const doc = new dom().parseFromString(documentContent, 'text/xml');
    const select = xpath.useNamespaces({ "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main" });
    const tags = JSON.parse(jsonContent);

    const processTagValue = (value) => {
        if (Array.isArray(value)) {
            return value.map((item, index) => {
                if (typeof item === 'string') {
                    return `${index + 1}. ${item}`;
                } else if (typeof item === 'object' && item.text) {
                    return `${index + 1}. ${item.highlighted ? item.text : item.text}`;
                }
                return '';
            }).join('\n');
        } else if (typeof value === 'object' && value.text) {
            return value.highlighted ? value.text : value.text;
        }
        return value;
    };

    select("//w:sdt", doc).forEach(node => {
        const tagNode = select('.//w:tag/@w:val', node)[0];
        if (tagNode) {
            const tagName = tagNode.value;
            if (tags.hasOwnProperty(tagName)) {
                const textNodes = select('.//w:t', node);
                if (textNodes.length > 0) {
                    const processedValue = processTagValue(tags[tagName]);
                    textNodes.forEach(textNode => {
                        textNode.textContent = processedValue;
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

const documentContent = readZipFile(zipFilePath, fileName);
const jsonContent = fs.readFileSync(jsonPath, 'utf8');

const updatedContent = replacePlaceholder(documentContent, jsonContent);
writeZipFile(zipFilePath, fileName, updatedContent);

console.log('Document updated successfully');

================
    ============

const fs = require('fs');
const AdmZip = require('adm-zip');
var xpath = require('xpath');
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
};

const replacePlaceholder = (documentContent, jsonContent) => {
    const doc = new dom().parseFromString(documentContent, 'text/xml');
    const select = xpath.useNamespaces({ "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main" });
    const tags = JSON.parse(jsonContent);

    const processTagValue = (value) => {
        if (Array.isArray(value)) {
            return value.map((item, index) => {
                if (typeof item === 'string') {
                    return `${index + 1}. ${item}`;
                } else if (typeof item === 'object' && item.text) {
                    return `${index + 1}. ${item.highlighted ? item.text : item.text}`;
                }
                return '';
            }).join('<w:br/>');
        } else if (typeof value === 'object' && value.text) {
            return value.highlighted ? value.text : value.text;
        }
        return value;
    };

    select("//w:sdt", doc).forEach(node => {
        const tagNode = select('.//w:tag/@w:val', node)[0];
        if (tagNode) {
            const tagName = tagNode.value;
            if (tags.hasOwnProperty(tagName)) {
                const textNodes = select('.//w:t', node);
                if (textNodes.length > 0) {
                    const processedValue = processTagValue(tags[tagName]);
                    textNodes.forEach((textNode, index) => {
                        if (index === 0) {
                            textNode.textContent = processedValue;
                        } else {
                            textNode.parentNode.removeChild(textNode);
                        }
                    });

                    // Adding line breaks for additional lines
                    if (textNodes.length > 0) {
                        const newTextNodes = processedValue.split('<w:br/>');
                        for (let i = 1; i < newTextNodes.length; i++) {
                            const newNode = doc.createElement('w:t');
                            const breakNode = doc.createElement('w:br');
                            newNode.textContent = newTextNodes[i];
                            textNodes[0].parentNode.appendChild(breakNode);
                            textNodes[0].parentNode.appendChild(newNode);
                        }
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

const documentContent = readZipFile(zipFilePath, fileName);
const jsonContent = fs.readFileSync(jsonPath, 'utf8');

const updatedContent = replacePlaceholder(documentContent, jsonContent);
writeZipFile(zipFilePath, fileName, updatedContent);

console.log('Document updated successfully');
==============

    const fs = require('fs');
const AdmZip = require('adm-zip');
var xpath = require('xpath');
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
};

const replacePlaceholder = (documentContent, jsonContent) => {
    const doc = new dom().parseFromString(documentContent, 'text/xml');
    const select = xpath.useNamespaces({ "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main" });
    const tags = JSON.parse(jsonContent);

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

    select("//w:sdt", doc).forEach(node => {
        const tagNode = select('.//w:tag/@w:val', node)[0];
        if (tagNode) {
            const tagName = tagNode.value;
            if (tags.hasOwnProperty(tagName)) {
                const textNodes = select('.//w:t', node);
                if (textNodes.length > 0) {
                    const processedValue = processTagValue(tags[tagName]);
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

    const serializer = new xmlSerializer();
    return serializer.serializeToString(doc);
};

const zipFilePath = './Test Document.docx';
const fileName = 'word/document.xml';
const jsonPath = './myoutput.json';

const documentContent = readZipFile(zipFilePath, fileName);
const jsonContent = fs.readFileSync(jsonPath, 'utf8');

const updatedContent = replacePlaceholder(documentContent, jsonContent);
writeZipFile(zipFilePath, fileName, updatedContent);

console.log('Document updated successfully');


=========


    const fs = require('fs');
const AdmZip = require('adm-zip');
var xpath = require('xpath');
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
};

const replacePlaceholder = (documentContent, jsonContent) => {
    const doc = new dom().parseFromString(documentContent, 'text/xml');
    const select = xpath.useNamespaces({ "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main" });
    const tags = JSON.parse(jsonContent);

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

    select("//w:sdt", doc).forEach(node => {
        const tagNode = select('.//w:tag/@w:val', node)[0];
        if (tagNode) {
            const tagName = tagNode.value;
            if (tags.hasOwnProperty(tagName)) {
                const textNodes = select('.//w:t', node);
                if (textNodes.length > 0) {
                    const processedValue = processTagValue(tags[tagName]);
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

    const serializer = new xmlSerializer();
    return serializer.serializeToString(doc);
};

const zipFilePath = './Test Document.docx';
const fileName = 'word/document.xml';
const jsonPath = './myoutput.json';

const documentContent = readZipFile(zipFilePath, fileName);
const jsonContent = fs.readFileSync(jsonPath, 'utf8');

const updatedContent = replacePlaceholder(documentContent, jsonContent);
writeZipFile(zipFilePath, fileName, updatedContent);

console.log('Document updated successfully');

