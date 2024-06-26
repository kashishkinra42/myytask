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
  function traverseContent(contentNode) {
    if (contentNode['w:t'] && Array.isArray(contentNode['w:t'])) {
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

const extractContentControlTags = async (filePath) => {
  try {
    const data = await readFileContent(filePath);
    const documentXml = await extractXmlFromZip(data);
    const xmlParsed = await parseXml(documentXml);

    const tags = {};
    const body = xmlParsed['w:document']['w:body'][0];
    traverseNodes(body, tags);

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































const express = require('express');
const multer = require('multer');
const { extractContentControlTags } = require('./extractTags');

const app = express();
const upload = multer({ dest: 'uploads/' });

const log = (message) => {
  console.log(`[${new Date().toISOString()}] ${message}`);
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
    res.status(200).json(tags);
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

=====

  const fs = require('fs');
const JSZip = require('jszip');
const xml2js = require('xml2js');

const log = (message) => {
  console.log(`[${new Date().toISOString()}] ${message}`);
};

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
  function traverseContent(contentNode) {
    if (contentNode['w:t'] && Array.isArray(contentNode['w:t'])) {
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

const extractContentControlTags = async (filePath) => {
  try {
    const data = await readFileContent(filePath);
    const documentXml = await extractXmlFromZip(data);
    const xmlParsed = await parseXml(documentXml);

    const tags = {};
    const body = xmlParsed['w:document']['w:body'][0];
    traverseNodes(body, tags);

    return tags;
  } catch (error) {
    console.error('Error:', error);
    throw error;
  }
};

module.exports = { extractContentControlTags };
