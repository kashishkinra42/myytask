const assert = require('assert');
const sinon = require('sinon');
const fs = require('fs');
const JSZip = require('jszip');
const xml2js = require('xml2js');
const proxyquire = require('proxyquire');
const path = require('path');

// Mock XML content for different scenarios
const validXmlContent = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:sdt>
      <w:sdtPr>
        <w:tag w:val="tag1"/>
      </w:sdtPr>
      <w:sdtContent>
        <w:p>
          <w:r>
            <w:t>Hello, World!</w:t>
          </w:r>
        </w:p>
      </w:sdtContent>
    </w:sdt>
  </w:body>
</w:document>`;

const emptyXmlContent = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
  </w:body>
</w:document>`;

const invalidXmlContent = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:sdt>
      <w:sdtPr>
        <w:tag w:val="tag1"/>
      </w:sdtPr>
      <w:sdtContent>
        <w:p>
          <w:r>
            <w:t>Hello, World!
          </w:r>
        </w:p>
      </w:sdtContent>
    </w:sdt>
  </w:body>
</w:document>`;

describe('Document Processing Functions', () => {
  let readFileStub;
  let jszipStub;
  let parseStringPromiseStub;

  const mockDocxFilePath = path.join(__dirname, 'updated_document.docx');
  const mockDocxFile = fs.readFileSync(mockDocxFilePath); // Read the binary content of the DOCX file

  beforeEach(() => {
    // Mock fs.promises.readFile
    readFileStub = sinon.stub(fs.promises, 'readFile').resolves(mockDocxFile);

    // Mock JSZip.loadAsync and .file().async()
    jszipStub = sinon.stub(JSZip, 'loadAsync').resolves({
      file: sinon.stub().returns({
        async: sinon.stub().resolves(validXmlContent),
      }),
    });

    // Mock xml2js.Parser().parseStringPromise
    parseStringPromiseStub = sinon.stub().resolves({
      'w:document': {
        'w:body': [
          {
            'w:sdt': [
              {
                'w:sdtPr': [
                  {
                    'w:tag': [
                      {
                        $: { 'w:val': 'tag1' },
                      },
                    ],
                  },
                ],
                'w:sdtContent': [
                  {
                    'w:p': [
                      {
                        'w:r': [
                          {
                            'w:t': ['Hello, World!'],
                          },
                        ],
                      },
                    ],
                  },
                ],
              },
            ],
          },
        ],
      },
    });

    // Use proxyquire to inject mocks
    this.functions = proxyquire('./upload.js', {
      fs,
      jszip: JSZip,
      xml2js: {
        Parser: function () {
          return { parseStringPromise: parseStringPromiseStub };
        },
      },
    });
  });

  afterEach(() => {
    sinon.restore();
  });

  describe('readFileContent', () => {
    it('should read file content successfully', async () => {
      const content = await this.functions.readFileContent(mockDocxFilePath);
      assert(content);
      assert(readFileStub.calledWith(mockDocxFilePath));
    });

    it('should handle file read errors', async () => {
      readFileStub.rejects(new Error('File not found'));
      await assert.rejects(this.functions.readFileContent('invalid_path.docx'), /File not found/);
    });
  });

  describe('extractXmlFromZip', () => {
    it('should extract XML content from zip successfully', async () => {
      const xmlContent = await this.functions.extractXmlFromZip(mockDocxFile);
      assert.strictEqual(xmlContent, validXmlContent);
    });

    it('should handle zip extraction errors', async () => {
      jszipStub.rejects(new Error('Invalid zip file'));
      await assert.rejects(this.functions.extractXmlFromZip(mockDocxFile), /Invalid zip file/);
    });
  });

  describe('parseXml', () => {
    it('should parse valid XML content successfully', async () => {
      const parsedXml = await this.functions.parseXml(validXmlContent);
      assert(parsedXml);
      assert(parseStringPromiseStub.calledWith(validXmlContent));
    });

    it('should handle invalid XML content', async () => {
      parseStringPromiseStub.rejects(new Error('Invalid XML'));
      await assert.rejects(this.functions.parseXml(invalidXmlContent), /Invalid XML/);
    });
  });

  describe('extractTagValue', () => {
    it('should extract tag value from content', () => {
      const content = {
        'w:p': [
          {
            'w:r': [
              {
                'w:t': ['Hello, World!'],
              },
            ],
          },
        ],
      };
      const result = this.functions.extractTagValue(content);
      assert.strictEqual(result, 'Hello, World!');
    });

    it('should handle empty content', () => {
      const content = {};
      const result = this.functions.extractTagValue(content);
      assert.strictEqual(result, '');
    });
  });

  describe('extractTagsFromBody', () => {
    it('should extract tags from body content', () => {
      const bodyContent = {
        'w:sdt': [
          {
            'w:sdtPr': [
              {
                'w:tag': [
                  {
                    $: { 'w:val': 'tag1' },
                  },
                ],
              },
            ],
            'w:sdtContent': [
              {
                'w:p': [
                  {
                    'w:r': [
                      {
                        'w:t': ['Hello, World!'],
                      },
                    ],
                  },
                ],
              },
            ],
          },
        ],
      };
      const tags = this.functions.extractTagsFromBody(bodyContent);
      assert.deepStrictEqual(tags, { tag1: 'Hello, World!' });
    });

    it('should handle body with no tags', () => {
      const bodyContent = {};
      const tags = this.functions.extractTagsFromBody(bodyContent);
      assert.deepStrictEqual(tags, {});
    });
  });

  describe('extractContentControlTags', () => {
    it('should extract content control tags from the document', async () => {
      const result = await this.functions.extractContentControlTags(mockDocxFilePath);
      assert.deepStrictEqual(result, { tag1: 'Hello, World!' });
    });

    it('should handle empty XML content', async () => {
      sinon.restore(); // Restore all stubs
      jszipStub = sinon.stub(JSZip, 'loadAsync').resolves({
        file: sinon.stub().returns({
          async: sinon.stub().resolves(emptyXmlContent),
        }),
      });

      const functions = proxyquire('./upload.js', {
        fs,
        jszip: JSZip,
        xml2js: {
          Parser: function () {
            return { parseStringPromise: parseStringPromiseStub };
          },
        },
      });

      const result = await functions.extractContentControlTags(mockDocxFilePath);
      assert.deepStrictEqual(result, {});
    });

    it('should handle invalid XML content', async () => {
      sinon.restore(); // Restore all stubs
      parseStringPromiseStub = sinon.stub().rejects(new Error('Invalid XML'));

      const functions = proxyquire('./upload.js', {
        fs,
        jszip: JSZip,
        xml2js: {
          Parser: function () {
            return { parseStringPromise: parseStringPromiseStub };
          },
        },
      });

      await assert.rejects(functions.extractContentControlTags(mockDocxFilePath), /Invalid XML/);
    });
  });
});
