app.js
  
const express = require('express');
const app = express();

app.use(express.json());

const log = (message) => {
  console.log(`[${new Date().toISOString()}] ${message}`);
};

app.post('/update-document', (req, res) => {
  log(`Received JSON: ${JSON.stringify(req.body)}`);
  const jsonContent = req.body;

  // Here, you could save the JSON content to a database or perform other operations if needed.

  res.status(200).json({ jsonContent });
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  log(`Server is running on port ${PORT}`);
});


postdata.js


const fs = require('fs');
const axios = require('axios');
const AdmZip = require('adm-zip');
const xpath = require('xpath');
const { DOMParser, XMLSerializer } = require('@xmldom/xmldom');
const { v4: uuidv4 } = require('uuid');

const log = (message) => {
  console.log(`[${new Date().toISOString()}] ${message}`);
};

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
  const doc = new DOMParser().parseFromString(documentContent, 'text/xml');
  const select = xpath.useNamespaces({ "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main" });

  select("//w:sdt", doc).forEach(node => {
    const tagNode = select('.//w:tag/@w:val', node)[0];
    if (tagNode) {
      const tagName = tagNode.value;
      const tagValue = jsonContent[tagName];
      
      if (tagValue !== null && tagValue !== undefined) {
        const textNodes = select('.//w:t', node);
        
        if (textNodes.length > 0) {
          textNodes.forEach(textNode => {
            textNode.textContent = tagValue;
          });
        }
      }
    }
  });
  
  const serializer = new XMLSerializer();
  const updatedContent = serializer.serializeToString(doc);
  return updatedContent;
};

// Read JSON content from output.json
const jsonData = JSON.parse(fs.readFileSync('myoutput.json', 'utf8'));

// Define the URL of your server endpoint
const url = 'http://localhost:3000/update-document';

// Post the JSON content to the server
axios.post(url, jsonData, {
  responseType: 'arraybuffer', // Ensure response type is handled as binary
  headers: {
    'Content-Type': 'application/json', // Set appropriate content type
  }
})
  .then(response => {
    log(`Status: ${response.status}`);
    log('Received response from server');

    const { data, headers } = response; // Destructure headers from response

    // Set the content disposition to inline to display the document in Postman
    const updatedFileName = `updated_${uuidv4()}.docx`;
    const contentDisposition = `attachment; filename="${updatedFileName}"`;

    // Output headers and data to console
    console.log(headers);
    console.log(data); // This should output the binary content of the updated document
  })
  .catch(error => {
    console.error('Error posting data:', error.message);
  });

    console.log(data); // This should output the binary content of the updated document
  })
  .catch(error => {
    console.error('Error posting data:', error.message);
  });
