const fs = require("fs");
const vm = require("vm");
const assert = require("assert");

const source = fs.readFileSync("script.js", "utf8");
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

const sample = `<?xml version="1.0" encoding="UTF-8"?>
<yml_catalog date="2026-10-09 15:00">
  <shop>
    <categories>
      <category id="1">Автотовары</category>
      <category id="2" parentId="1">Коврики</category>
      <category id="3" parentId="2"><![CDATA[Коврики EVA]]></category>
    </categories>
    <offers>
      <offer id="SKU-1">
        <categoryId>3</categoryId>
        <name>Коврик водительский</name>
      </offer>
      <offer id="SKU-2">
        <categoryId>2</categoryId>
        <name>Комплект ковриков</name>
      </offer>
      <offer id="SKU-3">
        <categoryId>999</categoryId>
        <name>Без категории</name>
      </offer>
    </offers>
  </shop>
</yml_catalog>`;

const result = sandbox.convertPromYmlToSupplementalXml_(sample);

assert.strictEqual(result.itemsTotal, 3);
assert.strictEqual(result.itemsWritten, 2);
assert.strictEqual(result.missingCategory, 1);
assert(result.xml.includes("<g:id>SKU-1</g:id>"));
assert(result.xml.includes("<g:product_type>Автотовары &gt; Коврики &gt; Коврики EVA</g:product_type>"));
assert(result.xml.includes("<g:id>SKU-2</g:id>"));
assert(result.xml.includes("<g:product_type>Автотовары &gt; Коврики</g:product_type>"));

console.log("OK");
