const fs = require('fs');
const JSZip = require('jszip');
const xml2js = require('xml2js');

async function extractContentControlTags(fpath) {
  try {
    const documentXml = await readDocumentXml(fpath);
    const tags = await parseDocumentXml(documentXml);
    saveTagsToFile(tags, 'myoutputk.json');
  } catch (error) {
    console.error('Error:', error);
  }
}

async function readDocumentXml(fpath) {
  const data = fs.readFileSync(fpath);
  const zip = await JSZip.loadAsync(data);
  const documentXml = await zip.file('word/document.xml').async('text');
  return documentXml;
}

async function parseDocumentXml(documentXml) {
  const parser = new xml2js.Parser();
  return new Promise((resolve, reject) => {
    parser.parseString(documentXml, (err, result) => {
      if (err) {
        return reject(err);
      }
      const tags = extractTagsFromBody(result['w:document']['w:body'][0]);
      resolve(tags);
    });
  });
}

function extractTagsFromBody(body) {
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
}

function extractTagValue(content) {
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
        if (rNode['w:rPr'] && rNode['w:rPr'][0]['w:highlight']) {
          highlighted = true;
        }
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

    if (highlighted) {
      return { text: nodeText, highlighted: true };
    }
    return nodeText;
  }

  traverseContent(content);

  if (listItems.length > 0) {
    // Return list items as an array of strings or objects
    return listItems;
  } else {
    return text;
  }
}

function saveTagsToFile(tags, outputPath) {
  const jsonFile = JSON.stringify(tags, null, 2);
  fs.writeFileSync(outputPath, jsonFile);
}

const fpath = './testttt.docx';
extractContentControlTags(fpath);
