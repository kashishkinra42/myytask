
const { get } = require('express/lib/response');
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
        // console.log(node);
        if (node['w:sdt']) {
          node['w:sdt'].forEach(sdt => {
            const tag = sdt['w:sdtPr'][0]['w:tag'];
            if (tag && tag[0]['$'] && tag[0]['$']['w:val']) {
              const tagName = tag[0]['$']['w:val'];
              const tagValue = extractValueOftag(sdt['w:sdtContent'][0]);
              if(tagValue.trim() !== ''){
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

      function extractValueOftag(content){
        let text ="";
        function traverseContent(contentNode){
          if(contentNode['w:t']){
            contentNode['w:t'].forEach(textNode => {
              if(typeof textNode === 'string'){
                text += textNode;
              }else if(textNode['_']){
                text += textNode['_'];
              }
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
      traverseNodes(body);

      // console.log(tags);
      const jsonFile = JSON.stringify(tags, null, 2);
      fs.writeFileSync('myoutput.json', jsonFile);
    });
  }
  catch (error) {
    console.error('Error : ', error);
  }
}

const fpath = './Test Document.docx';
extractContentControlTags(fpath);






----------




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
              const tagValue = extractValueOfTag(sdt['w:sdtContent'][0]);
              if(tagValue.trim() !== ''){
                tags[tagName] = tagValue;
              }
            }
          });
        }
        if (node['w:p']) {
          node['w:p'].forEach(p => {
            if (p['w:numPr'] && p['w:numPr'][0]['w:numId']) {
              const numId = p['w:numPr'][0]['w:numId'][0]['$']['w:val'];
              const listValue = extractNumberedListValue(p, numId);
              tags[`numid_${numId}`] = listValue;
            }
          });
        }
        Object.values(node).forEach(value => {
          if (Array.isArray(value)) {
            value.forEach(child => traverseNodes(child));
          }
        });
      }

      function extractValueOfTag(content) {
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

          Object.values(contentNode).forEach(value => {
            if (Array.isArray(value)) {
              value.forEach(child => traverseContent(child));
            }
          });
        }
        traverseContent(content);
        return text;
      }

      function extractNumberedListValue(paragraph, numId) {
        let text = "";
        function traverseParagraph(p) {
          if (p['w:r']) {
            p['w:r'].forEach(run => {
              if (run['w:t']) {
                run['w:t'].forEach(t => {
                  if (typeof t === 'string') {
                    text += t;
                  } else if (t['_']) {
                    text += t['_'];
                  }
                });
              }
            });
          }
        }
        traverseParagraph(paragraph);
        return text.trim() ? `${numId}. ${text}` : '';
      }

      traverseNodes(body);

      const jsonFile = JSON.stringify(tags, null, 2);
      fs.writeFileSync('myoutput.json', jsonFile);
    });
  }
  catch (error) {
    console.error('Error : ', error);
  }
}

const fpath = './Test Document.docx';
extractContentControlTags(fpath);



---------

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
              const tagValue = extractValueOfTag(sdt['w:sdtContent'][0]);
              if(tagValue.trim() !== ''){
                tags[tagName] = tagValue;
              }
            }
          });
        }
        if (node['w:p']) {
          node['w:p'].forEach(p => {
            if (p['w:numPr'] && p['w:numPr'][0]['w:numId']) {
              const numId = p['w:numPr'][0]['w:numId'][0]['$']['w:val'];
              const listValue = extractNumberedListValue(p, numId);
              if (listValue.trim() !== '') {
                if (!tags[`numid_${numId}`]) {
                  tags[`numid_${numId}`] = [];
                }
                tags[`numid_${numId}`].push(listValue);
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

      function extractValueOfTag(content) {
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

          Object.values(contentNode).forEach(value => {
            if (Array.isArray(value)) {
              value.forEach(child => traverseContent(child));
            }
          });
        }
        traverseContent(content);
        return text;
      }

      function extractNumberedListValue(paragraph, numId) {
        let text = "";
        function traverseParagraph(p) {
          if (p['w:r']) {
            p['w:r'].forEach(run => {
              if (run['w:t']) {
                run['w:t'].forEach(t => {
                  if (typeof t === 'string') {
                    text += t;
                  } else if (t['_']) {
                    text += t['_'];
                  }
                });
              }
            });
          }
        }
        traverseParagraph(paragraph);
        return text.trim() ? `${text}` : '';
      }

      traverseNodes(body);

      // Format the numbered lists correctly
      for (const key in tags) {
        if (key.startsWith('numid_')) {
          tags[key] = tags[key].map((item, index) => `${index + 1}. ${item}`).join('\n');
        }
      }

      const jsonFile = JSON.stringify(tags, null, 2);
      fs.writeFileSync('myoutput.json', jsonFile);
    });
  }
  catch (error) {
    console.error('Error : ', error);
  }
}

const fpath = './Test Document.docx';
extractContentControlTags(fpath);





--------


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
              const tagValue = extractValueOfTag(sdt['w:sdtContent'][0]);
              if(tagValue.trim() !== ''){
                tags[tagName] = tagValue;
              }
            }
          });
        }
        if (node['w:p']) {
          node['w:p'].forEach(p => {
            if (p['w:numPr'] && p['w:numPr'][0]['w:numId']) {
              const numId = p['w:numPr'][0]['w:numId'][0]['$']['w:val'];
              const listValue = extractNumberedListValue(p);
              if (listValue.trim() !== '') {
                if (!tags[`numid_${numId}`]) {
                  tags[`numid_${numId}`] = [];
                }
                tags[`numid_${numId}`].push(listValue);
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

      function extractValueOfTag(content) {
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

          Object.values(contentNode).forEach(value => {
            if (Array.isArray(value)) {
              value.forEach(child => traverseContent(child));
            }
          });
        }
        traverseContent(content);
        return text;
      }

      function extractNumberedListValue(paragraph) {
        let text = "";
        function traverseParagraph(p) {
          if (p['w:r']) {
            p['w:r'].forEach(run => {
              if (run['w:t']) {
                run['w:t'].forEach(t => {
                  if (typeof t === 'string') {
                    text += t;
                  } else if (t['_']) {
                    text += t['_'];
                  }
                });
              }
            });
          }
        }
        traverseParagraph(paragraph);
        return text.trim();
      }

      traverseNodes(body);

      // Format the numbered lists correctly
      for (const key in tags) {
        if (key.startsWith('numid_')) {
          tags[key] = tags[key].map((item, index) => `${index + 1}. ${item}`).join('\n');
        }
      }

      const jsonFile = JSON.stringify(tags, null, 2);
      fs.writeFileSync('myoutput.json', jsonFile);
    });
  }
  catch (error) {
    console.error('Error : ', error);
  }
}

const fpath = './Test Document.docx';
extractContentControlTags(fpath);

