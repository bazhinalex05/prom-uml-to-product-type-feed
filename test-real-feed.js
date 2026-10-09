const fs = require("fs");
const vm = require("vm");

const source = fs.readFileSync("script.js", "utf8");
const xml = fs.readFileSync(process.argv[2], "utf8");
const sandbox = {
  Logger: { log() {} },
  Utilities: {},
  DriveApp: {},
  SpreadsheetApp: {},
  ContentService: {},
  MimeType: { XML: "application/xml" }
};

vm.createContext(sandbox);
vm.runInContext(source, sandbox);

const result = sandbox.convertPromYmlToSupplementalXml_(xml);
fs.writeFileSync("sample-elegancehome-product-type-supplemental.xml", result.xml, "utf8");

console.log(JSON.stringify({
  itemsTotal: result.itemsTotal,
  itemsWritten: result.itemsWritten,
  missingCategory: result.missingCategory,
  outputFile: "sample-elegancehome-product-type-supplemental.xml"
}, null, 2));
