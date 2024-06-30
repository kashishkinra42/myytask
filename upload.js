const fs = require('fs');
const JSZip = require('jszip');
const xml2js = require('xml2js');
const path = require('path');

// const log = (message) => {
//   console.log(`[${new Date().toISOString()}] ${message}`);
// };

const readFileContent = async (filePath) => {
  return fs.promises.readFile(filePath);
};

const extractXmlFromZip = async (data) => {
  const zip = await JSZip.loadAsync(data);
  return zip.file('word/document.xml').async('text');
};

const parseXml = async (xmlContent) => {
  const parser = new xml2js.Parser();
  return parser.parseStringPromise(xmlContent);
};

const traverseNodes = (node, tags) => {
  if (node['w:sdt']) {
    node['w:sdt'].forEach(sdt => {
      const sdtPr = sdt['w:sdtPr'] && Array.isArray(sdt['w:sdtPr']) ? sdt['w:sdtPr'][0] : null;
      const tag = sdtPr && sdtPr['w:tag'] && Array.isArray(sdtPr['w:tag']) ? sdtPr['w:tag'][0] : null;
      if (tag && tag['$'] && tag['$']['w:val']) {
        const tagName = tag['$']['w:val'];
        const sdtContent = sdt['w:sdtContent'] && Array.isArray(sdt['w:sdtContent']) ? sdt['w:sdtContent'][0] : null;
        const tagValue = sdtContent ? extractTagValue(sdtContent) : '';
        if (tagValue.trim() !== '') {
          tags[tagName] = tagValue;
        }
      }
    });
  }
  Object.values(node).forEach(value => {
    if (Array.isArray(value)) {
      value.forEach(child => traverseNodes(child, tags));
    }
  });
};

const extractTagValue = (content) => {
  let text = '';
  let listItems = [];

  function traverseContent(contentNode) {
    if (contentNode['w:p']) {
      contentNode['w:p'].forEach(pNode => {
        if (pNode['w:pPr'] && pNode['w:pPr'][0]['w:numPr']) {
          // Detect a numbered list item
          const listItem = extractTextFromNode(pNode);
          if (typeof listItem === 'string' && listItem.trim() !== '') {
            listItems.push(listItem);
          } else if (typeof listItem === 'object' && listItem.text.trim() !== '') {
            listItems.push(listItem);
          }
        } else {
          // Regular text node
          const pText = extractTextFromNode(pNode);
          if (typeof pText === 'string' && pText.trim() !== '') {
            text += pText;
          } else if (typeof pText === 'object' && pText.text.trim() !== '') {
            text += pText.text;
          }
        }
      });
    }
    Object.values(contentNode).forEach(value => {
      if (Array.isArray(value)) {
        value.forEach(child => traverseContent(child));
      }
    });
  }

  function extractTextFromNode(node) {
    let nodeText = '';
    let highlighted = false;

    if (node['w:r']) {
      node['w:r'].forEach(rNode => {
        if (rNode['w:t']) {
          rNode['w:t'].forEach(textNode => {
            if (typeof textNode === 'string') {
              nodeText += textNode;
            } else if (textNode['_']) {
              nodeText += textNode['_'];
            }
          });
        }
      });
    }
    return nodeText;
  }

  traverseContent(content);

  if (listItems.length > 0) {
    return listItems;
  } else {
    return text;
  }
};

const extractTagsFromBody = (body) => {
  const tags = {};

  function traverseNodes(node) {
    if (node['w:sdt']) {
      node['w:sdt'].forEach(sdt => {
        const tag = sdt['w:sdtPr'][0]['w:tag'];
        if (tag && tag[0]['$'] && tag[0]['$']['w:val']) {
          const tagName = tag[0]['$']['w:val'];
          const tagValue = extractTagValue(sdt['w:sdtContent'][0]);
          if (typeof tagValue === 'string' && tagValue.trim() !== '') {
            tags[tagName] = tagValue;
          } else if (Array.isArray(tagValue) && tagValue.length > 0) {
            tags[tagName] = tagValue;
          }
        }
      });
    }
    Object.values(node).forEach(value => {
      if (Array.isArray(value)) {
        value.forEach(child => traverseNodes(child));
      }
    });
  }

  traverseNodes(body);
  return tags;
};

const extractContentControlTags = async (filePath) => {
  try {
    const data = await readFileContent(filePath);
    const documentXml = await extractXmlFromZip(data);
    const xmlParsed = await parseXml(documentXml);

    const tags = extractTagsFromBody(xmlParsed['w:document']['w:body'][0]);

    return tags;
  } catch (error) {
    console.error('Error:', error);
    throw error;
  }
};

module.exports = { extractContentControlTags };
