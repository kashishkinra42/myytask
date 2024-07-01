Z:\Desktop\work-addin>npx mocha uploadtest.js


  Document Processing Functions
    readFileContent
      √ should read file content successfully
      √ should handle file read errors
    extractXmlFromZip
      √ should extract XML content from zip successfully
      √ should handle zip extraction errors
    parseXml
      √ should parse valid XML content successfully
      √ should handle invalid XML content
    extractTagValue
      √ should extract tag value from content
      √ should handle empty content
    extractTagsFromBody
      √ should extract tags from body content
      √ should handle body with no tags
    extractContentControlTags
      √ should extract content control tags from the document
      1) should handle empty XML content
      2) should handle invalid XML content


  11 passing (715ms)
  2 failing

  1) Document Processing Functions
       extractContentControlTags
         should handle empty XML content:

      AssertionError [ERR_ASSERTION]: Expected values to be strictly deep-equal:
+ actual - expected

+ {
+   tag1: 'Hello, World!'
+ }
- {}
      + expected - actual

      -{
      -  "tag1": "Hello, World!"
      -}
      +{}

      at Context.<anonymous> (uploadtest.js:242:14)

  2) Document Processing Functions
       extractContentControlTags
         should handle invalid XML content:
     TypeError: parseStringPromiseStub.restore is not a function
      at Context.<anonymous> (uploadtest.js:246:30)
      at process.processImmediate (node:internal/timers:478:21)
