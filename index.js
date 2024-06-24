
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

========================
  ============================
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

  
  ==========================================================
