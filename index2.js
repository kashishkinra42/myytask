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
const JSZip = require('jszip');
const xml2js = require('xml2js');

async function extractContentControlTags(fpath) {
  try {
    const data = fs.readFileSync(fpath);
    const zip = await JSZip.loadAsync(data);
    const documentXml = await zip.file('word/document.xml').async('text');
    const parser = new xml2js.Parser();

    parser.parseString(documentXml, (err, result) => {
      if (err) {
        throw err;
      }

      const tags = {};
      const body = result['w:document']['w:body'][0];

      function traverseNodes(node) {
        if (node['w:sdt']) {
          node['w:sdt'].forEach(sdt => {
            const tag = sdt['w:sdtPr'][0]['w:tag'];
            if (tag && tag[0]['$'] && tag[0]['$']['w:val']) {
              const tagName = tag[0]['$']['w:val'];
              const tagValue = extractValueOftag(sdt['w:sdtContent'][0]);
              if (tagValue.trim() !== '') {
                tags[tagName] = tagValue;
              }
            }
          });
        }
        if (node['w:numPr']) {
          tags['numList'] = extractNumberedList(node);
        }
        Object.values(node).forEach(value => {
          if (Array.isArray(value)) {
            value.forEach(child => traverseNodes(child));
          }
        });
      }

      function extractValueOftag(content) {
        let text = "";
        function traverseContent(contentNode) {
          if (contentNode['w:t']) {
            contentNode['w:t'].forEach(textNode => {
              if (typeof textNode === 'string') {
                text += textNode;
              } else if (textNode['_']) {
                text += textNode['_'];
              }
            });
          }
          if (contentNode['w:tc']) {
            contentNode['w:tc'].forEach(tcNode => {
              traverseContent(tcNode);
            });
          }
          Object.values(contentNode).forEach(value => {
            if (Array.isArray(value)) {
              value.forEach(child => traverseContent(child));
            }
          });
        }
        traverseContent(content);
        return text;
      }

      function extractNumberedList(node) {
        let list = [];
        function traverseList(contentNode, level = 0) {
          if (contentNode['w:numPr']) {
            contentNode['w:p'].forEach(pNode => {
              let text = "";
              if (pNode['w:r']) {
                pNode['w:r'].forEach(rNode => {
                  if (rNode['w:t']) {
                    text += rNode['w:t'].map(tNode => tNode._).join('');
                  }
                });
              }
              list.push({ level: level, text: text });
            });
          }
          Object.values(contentNode).forEach(value => {
            if (Array.isArray(value)) {
              value.forEach(child => traverseList(child, level + 1));
            }
          });
        }
        traverseList(node);
        return list;
      }

      traverseNodes(body);

      // Convert list to structured text with new lines
      if (tags['numList']) {
        let numberedText = "";
        tags['numList'].forEach(item => {
          numberedText += `${'  '.repeat(item.level)}${item.text}\n`;
        });
        tags['numList'] = numberedText.trim();
      }

      const jsonFile = JSON.stringify(tags, null, 2);
      fs.writeFileSync('myoutput.json', jsonFile);
    });
  } catch (error) {
    console.error('Error:', error);
  }
}

const fpath = './Test Document.docx';
extractContentControlTags(fpath);
