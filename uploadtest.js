const fs = require('fs');
const JSZip = require('jszip');
const xml2js = require('xml2js');
const { extractContentControlTags } = require('./path_to_your_module');

// Mock fs.promises.readFile
jest.mock('fs', () => ({
  promises: {
    readFile: jest.fn(),
  },
}));

// Mock data for the ZIP file
const mockDocxFile = 'path_to_your_mock_docx_file';
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

describe('extractContentControlTags', () => {
  beforeEach(() => {
    fs.promises.readFile.mockClear();
  });

  it('should extract content control tags from the document', async () => {
    fs.promises.readFile.mockResolvedValue(mockDocxFile);
    
    // Mock JSZip.loadAsync to return the mock XML content
    JSZip.loadAsync = jest.fn().mockResolvedValue({
      file: jest.fn().mockReturnValue({
        async: jest.fn().mockResolvedValue(mockXmlContent),
      }),
    });

    // Mock xml2js.Parser().parseStringPromise to parse the mock XML content
    const parseStringPromise = jest.fn().mockResolvedValue({
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

    xml2js.Parser = jest.fn().mockImplementation(() => ({
      parseStringPromise,
    }));

    const result = await extractContentControlTags('path_to_mock_file.docx');
    
    expect(result).toEqual({
      tag1: 'Hello, World!',
    });

    expect(fs.promises.readFile).toHaveBeenCalledWith('path_to_mock_file.docx');
    expect(JSZip.loadAsync).toHaveBeenCalledWith(mockDocxFile);
    expect(parseStringPromise).toHaveBeenCalledWith(mockXmlContent);
  });

  it('should handle errors gracefully', async () => {
    fs.promises.readFile.mockRejectedValue(new Error('File not found'));

    await expect(extractContentControlTags('invalid_path.docx')).rejects.toThrow('File not found');
  });
});
