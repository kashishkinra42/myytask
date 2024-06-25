const express = require('express');
const fs = require('fs');
const JSZip = require('jszip');
const xml2js = require('xml2js');
const { v4: uuidv4 } = require('uuid');
const multer = require('multer');

const app = express();
const upload = multer({ dest: 'uploads/' });

const log = (message) => {
  console.log(`[${new Date().toISOString()}] ${message}`);
};

const readDocumentXml = async (fpath) => {
  const data = fs.readFileSync(fpath);
  const zip = await JSZip.loadAsync(data);
  const documentXml = await zip.file('word/document.xml').async('text');
  return documentXml;
};

const parseDocumentXml = (documentXml) => {
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

  traverseNodes(body);
  return tags;
};

const extractTagValue = (content) => {
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
};

app.post('/extract-tags', upload.single('document'), async (req, res) => {
  log('Received file upload request');

  try {
    const filePath = req.file.path;
    const documentXml = await readDocumentXml(filePath);
    const tags = await parseDocumentXml(documentXml);

    fs.unlinkSync(filePath); // Clean up the uploaded file

    res.status(200).json(tags);
  } catch (error) {
    console.error('Error extracting tags:', error);
    res.status(500).send('Internal Server Error');
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  log(`Server is running on port ${PORT}`);
});
