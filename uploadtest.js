const assert = require('assert');
const sinon = require('sinon');
const fs = require('fs');
const JSZip = require('jszip');
const xml2js = require('xml2js');
const proxyquire = require('proxyquire');
const path = require('path');

describe('extractContentControlTags', () => {
  let readFileStub;
  let jszipStub;
  let parseStringPromiseStub;

  const mockDocxFilePath = path.join(__dirname, 'mock_document.docx');
  const mockDocxFile = fs.readFileSync(mockDocxFilePath); // Read the binary content of the DOCX file

  beforeEach(() => {
    // Mock fs.promises.readFile
    readFileStub = sinon.stub(fs.promises, 'readFile').resolves(mockDocxFile);

    // Mock JSZip.loadAsync and .file().async()
    jszipStub = sinon.stub(JSZip, 'loadAsync').resolves({
      file: sinon.stub().returns({
        async: sinon.stub().resolves(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
        </w:document>`),
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
    const { extractContentControlTags } = proxyquire('../path_to_your_module', {
      fs,
      jszip: JSZip,
      xml2js: {
        Parser: function () {
          return { parseStringPromise: parseStringPromiseStub };
        },
      },
    });

    this.extractContentControlTags = extractContentControlTags;
  });

  afterEach(() => {
    sinon.restore();
  });

  it('should extract content control tags from the document', async () => {
    const result = await this.extractContentControlTags(mockDocxFilePath);

    assert.deepEqual(result, {
      tag1: 'Hello, World!',
    });

    assert(readFileStub.calledWith(mockDocxFilePath));
    assert(jszipStub.calledWith(mockDocxFile));
    assert(parseStringPromiseStub.calledOnce);
  });

  it('should handle errors gracefully', async () => {
    readFileStub.rejects(new Error('File not found'));

    await assert.rejects(async () => {
      await this.extractContentControlTags('invalid_path.docx');
    }, new Error('File not found'));
  });
});


=========


  const assert = require('assert');
const sinon = require('sinon');
const fs = require('fs');
const JSZip = require('jszip');
const xml2js = require('xml2js');
const proxyquire = require('proxyquire');
const path = require('path');

describe('extractContentControlTags', () => {
  let readFileStub;
  let jszipStub;
  let parseStringPromiseStub;

  const mockDocxFilePath = path.join(__dirname, 'mock_document.docx');
  const mockDocxFile = fs.readFileSync(mockDocxFilePath); // Read the binary content of the DOCX file

  const mockXmlContent = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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

  beforeEach(() => {
    // Mock fs.promises.readFile
    readFileStub = sinon.stub(fs.promises, 'readFile').resolves(mockDocxFile);

    // Mock JSZip.loadAsync and .file().async()
    jszipStub = sinon.stub(JSZip, 'loadAsync').resolves({
      file: sinon.stub().returns({
        async: sinon.stub().resolves(mockXmlContent),
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
    const { extractContentControlTags } = proxyquire('./path_to_your_module', {
      fs,
      jszip: JSZip,
      xml2js: {
        Parser: function () {
          return { parseStringPromise: parseStringPromiseStub };
        },
      },
    });

    this.extractContentControlTags = extractContentControlTags;
  });

  afterEach(() => {
    sinon.restore();
  });

  it('should extract content control tags from the document', async () => {
    const result = await this.extractContentControlTags('path_to_mock_file.docx');

    assert.deepEqual(result, {
      tag1: 'Hello, World!',
    });

    assert(readFileStub.calledWith('path_to_mock_file.docx'));
    assert(jszipStub.calledWith(mockDocxFile));
    assert(parseStringPromiseStub.calledWith(mockXmlContent));
  });

  it('should handle errors gracefully', async () => {
    readFileStub.rejects(new Error('File not found'));

    await assert.rejects(async () => {
      await this.extractContentControlTags('invalid_path.docx');
    }, new Error('File not found'));
  });
});
