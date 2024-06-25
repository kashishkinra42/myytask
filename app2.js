upload.js

const fs = require('fs');
const path = require('path');
const axios = require('axios');
const FormData = require('form-data');

// Replace with the path to your Word document
const filePath = path.join(__dirname, 'your-document.docx');
const url = 'http://localhost:3000/extract-tags';

async function uploadDocument() {
  try {
    const form = new FormData();
    form.append('file', fs.createReadStream(filePath));

    const response = await axios.post(url, form, {
      headers: {
        ...form.getHeaders(),
      },
    });

    console.log(`Status: ${response.status}`);
    console.log('Response:', response.data);
  } catch (error) {
    console.error('Error uploading document:', error.message);
  }
}

uploadDocument();



----------------------------------------------------------



  app2.js
const express = require('express');
const fs = require('fs');
const multer = require('multer');
const JSZip = require('jszip');
const xml2js = require('xml2js');

const app = express();
const upload = multer({ dest: 'uploads/' });

const log = (message) => {
  console.log(`[${new Date().toISOString()}] ${message}`);
};

const extractContentControlTags = async (filePath) => {
  try {
    const data = fs.readFileSync(filePath);
    const zip = await JSZip.loadAsync(data);
    const documentXml = await zip.file('word/document.xml').async('text');
    const parser = new xml2js.Parser();
    const tags = {};

    await parser.parseStringPromise(documentXml).then((result) => {
      const body = result['w:document']['w:body'][0];

      function traverseNodes(node) {
        if (node['w:sdt']) {
          node['w:sdt'].forEach(sdt => {
            const tag = sdt['w:sdtPr'][0]['w:tag'];
            if (tag && tag[0]['$'] && tag[0]['$']['w:val']) {
              const tagName = tag[0]['$']['w:val'];
              const tagValue = extractTagValue(sdt['w:sdtContent'][0]);
              if (tagValue.trim() !== '') {
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

      function extractTagValue(content) {
        let text = '';
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
      traverseNodes(body);
    });

    return tags;
  } catch (error) {
    console.error('Error:', error);
    throw error;
  }
};

app.post('/extract-tags', upload.single('file'), async (req, res) => {
  if (!req.file) {
    log(`No file received in the request`);
    return res.status(400).send('No file uploaded');
  }

  log(`Received file for extraction: ${req.file.path}`);
  const filePath = req.file.path;

  try {
    const tags = await extractContentControlTags(filePath);
    const jsonFilePath = './myoutput.json';
    fs.writeFileSync(jsonFilePath, JSON.stringify(tags, null, 2));
    res.status(200).send(`Tags extracted successfully. Saved as ${jsonFilePath}`);
  } catch (error) {
    console.error('Error extracting tags:', error);
    res.status(500).send('Internal Server Error');
  } finally {
    fs.unlinkSync(filePath); // Clean up the uploaded file
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  log(`Server is running on port ${PORT}`);
});

