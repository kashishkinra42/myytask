const express = require('express');
const fs = require('fs');
const AdmZip = require('adm-zip');
const xpath = require('xpath');
const { DOMParser, XMLSerializer } = require('@xmldom/xmldom');

const app = express();
app.use(express.json());

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
    const doc = new DOMParser().parseFromString(documentContent, 'text/xml');
    const select = xpath.useNamespaces({ "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main" });
    const tags = jsonContent;

    select("//w:sdt", doc).forEach(node => {
        var tagNode = select('.//w:tag/@w:val', node)[0];
        if (tagNode) {
            const tagName = tagNode.value;
            if (tags.hasOwnProperty(tagName)) {
                const textNodes = select('.//w:t', node);
                if (textNodes.length > 0) {
                    textNodes.forEach(textNode => {
                        textNode.textContent = tags[tagName];
                    });
                }
            }
        }
    });
    const serializer = new XMLSerializer();
    return serializer.serializeToString(doc);
};

app.post('/update-document', (req, res) => {
    const zipFilePath = './Test Document.docx';
    const fileName = 'word/document.xml';
    const jsonContent = req.body;

    try {
        const documentContent = readZipFile(zipFilePath, fileName);
        const updatedContent = replacePlaceholder(documentContent, jsonContent);
        writeZipFile(zipFilePath, fileName, updatedContent);

        res.status(200).send('Document updated successfully');
    } catch (error) {
        console.error('Error updating document:', error);
        res.status(500).send('Internal Server Error');
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(Server is running on port ${PORT});
});
