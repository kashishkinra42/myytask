Z:\Desktop\work-addin>npx mocha uploadtest.js


  Document Processing Functions
    readFileContent
      1) should read file content successfully
      2) should handle file read errors
    extractXmlFromZip
      3) should extract XML content from zip successfully
      4) should handle zip extraction errors
    parseXml
      5) should parse valid XML content successfully
      6) should handle invalid XML content
    extractTagValue
      7) should extract tag value from content
      8) should handle empty content
    extractTagsFromBody
      9) should extract tags from body content
      10) should handle body with no tags
    extractContentControlTags
      √ should extract content control tags from the document
      11) should handle empty XML content
      12) should handle invalid XML content


  1 passing (706ms)
  12 failing

  1) Document Processing Functions
       readFileContent
         should read file content successfully:
     TypeError: this.functions.readFileContent is not a function
      at Context.<anonymous> (uploadtest.js:124:44)
      at process.processImmediate (node:internal/timers:478:21)

  2) Document Processing Functions
       readFileContent
         should handle file read errors:
     TypeError: this.functions.readFileContent is not a function
      at Context.<anonymous> (uploadtest.js:131:43)
      at process.processImmediate (node:internal/timers:478:21)

  3) Document Processing Functions
       extractXmlFromZip
         should extract XML content from zip successfully:
     TypeError: this.functions.extractXmlFromZip is not a function
      at Context.<anonymous> (uploadtest.js:137:47)
      at process.processImmediate (node:internal/timers:478:21)

  4) Document Processing Functions
       extractXmlFromZip
         should handle zip extraction errors:
     TypeError: this.functions.extractXmlFromZip is not a function
      at Context.<anonymous> (uploadtest.js:143:43)
      at process.processImmediate (node:internal/timers:478:21)

  5) Document Processing Functions
       parseXml
         should parse valid XML content successfully:
     TypeError: this.functions.parseXml is not a function
      at Context.<anonymous> (uploadtest.js:149:46)
      at process.processImmediate (node:internal/timers:478:21)

  6) Document Processing Functions
       parseXml
         should handle invalid XML content:
     TypeError: this.functions.parseXml is not a function
      at Context.<anonymous> (uploadtest.js:156:43)
      at process.processImmediate (node:internal/timers:478:21)

  7) Document Processing Functions
       extractTagValue
         should extract tag value from content:
     TypeError: this.functions.extractTagValue is not a function
      at Context.<anonymous> (uploadtest.js:173:37)
      at process.processImmediate (node:internal/timers:478:21)

  8) Document Processing Functions
       extractTagValue
         should handle empty content:
     TypeError: this.functions.extractTagValue is not a function
      at Context.<anonymous> (uploadtest.js:179:37)
      at process.processImmediate (node:internal/timers:478:21)

  9) Document Processing Functions
       extractTagsFromBody
         should extract tags from body content:
     TypeError: this.functions.extractTagsFromBody is not a function
      at Context.<anonymous> (uploadtest.js:214:35)
      at process.processImmediate (node:internal/timers:478:21)

  10) Document Processing Functions
       extractTagsFromBody
         should handle body with no tags:
     TypeError: this.functions.extractTagsFromBody is not a function
      at Context.<anonymous> (uploadtest.js:220:35)
      at process.processImmediate (node:internal/timers:478:21)

  11) Document Processing Functions
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

      at Context.<anonymous> (uploadtest.js:240:14)

  12) Document Processing Functions
       extractContentControlTags
         should handle invalid XML content:
     TypeError: parseStringPromiseStub.restore is not a function
      at Context.<anonymous> (uploadtest.js:244:30)
      at process.processImmediate (node:internal/timers:478:21)



