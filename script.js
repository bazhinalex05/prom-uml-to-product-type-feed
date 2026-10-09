var CONTROL_SPREADSHEET_URL = "PASTE_CONTROL_SPREADSHEET_URL_HERE";
var CONFIG_SHEET_NAME = "Feeds";
var OUTPUT_FOLDER_NAME = "prom-product-type-feeds";
var WEB_APP_BASE_URL = "";
var SHARE_OUTPUT_FILES = true;
var MAX_FEEDS_PER_RUN = 20;

var COLS = [
  "enabled",
  "site",
  "slug",
  "source_feed_url",
  "output_feed_url",
  "drive_file_id",
  "last_status",
  "last_updated",
  "items_total",
  "items_written",
  "missing_category",
  "error"
];

function setupControlSheet() {
  var sheet = getControlSheet_();
  ensureHeader_(sheet);
  if (sheet.getLastRow() < 2) {
    sheet.getRange(2, 1, 1, COLS.length).setValues([[
      true,
      "example.com",
      "example",
      "https://example.com/prom-uml.yml",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      ""
    ]]);
    sheet.getRange(2, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
  }
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, COLS.length);
}

function runAllFeeds() {
  var sheet = getControlSheet_();
  ensureHeader_(sheet);
  var rows = getConfigRows_(sheet);
  var processed = 0;

  for (var i = 0; i < rows.length; i++) {
    if (processed >= MAX_FEEDS_PER_RUN) break;
    if (rows[i].enabled !== true) continue;
    processConfigRow_(sheet, rows[i]);
    processed++;
  }

  Logger.log("Feeds processed: " + processed);
}

function runSelectedRow(rowNumber) {
  var sheet = getControlSheet_();
  ensureHeader_(sheet);
  var row = Number(rowNumber || sheet.getActiveRange().getRow());
  if (row < 2) throw new Error("Select a data row, not the header.");
  var rows = getConfigRows_(sheet);
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].rowNumber === row) {
      processConfigRow_(sheet, rows[i]);
      return;
    }
  }
  throw new Error("Row not found: " + row);
}

function doGet(e) {
  var slug = String((e && e.parameter && e.parameter.slug) || "").trim();
  if (!slug) {
    return ContentService.createTextOutput("Missing slug").setMimeType(ContentService.MimeType.TEXT);
  }

  var sheet = getControlSheet_();
  ensureHeader_(sheet);
  var rows = getConfigRows_(sheet);
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].slug === slug && rows[i].drive_file_id) {
      var file = DriveApp.getFileById(rows[i].drive_file_id);
      return ContentService
        .createTextOutput(file.getBlob().getDataAsString("UTF-8"))
        .setMimeType(ContentService.MimeType.XML);
    }
  }

  return ContentService.createTextOutput("Feed not found").setMimeType(ContentService.MimeType.TEXT);
}

function processConfigRow_(sheet, row) {
  var now = new Date();
  try {
    validateRow_(row);
    writeStatus_(sheet, row.rowNumber, "RUNNING", now, "", "", "", "");

    var xmlText = fetchText_(row.source_feed_url);
    var conversion = convertPromYmlToSupplementalXml_(xmlText);
    if (!conversion.itemsWritten) {
      throw new Error("No items with matched category paths were found.");
    }

    var file = upsertOutputFile_(row, conversion.xml);
    if (SHARE_OUTPUT_FILES) {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    }

    var outputUrl = buildOutputUrl_(row.slug, file.getId());
    sheet.getRange(row.rowNumber, colIndex_("output_feed_url")).setValue(outputUrl);
    sheet.getRange(row.rowNumber, colIndex_("drive_file_id")).setValue(file.getId());
    writeStatus_(
      sheet,
      row.rowNumber,
      "OK",
      now,
      conversion.itemsTotal,
      conversion.itemsWritten,
      conversion.missingCategory,
      ""
    );
    Logger.log(row.slug + ": OK, items written " + conversion.itemsWritten);
  } catch (err) {
    writeStatus_(sheet, row.rowNumber, "ERROR", now, "", "", "", String(err && err.message ? err.message : err));
    Logger.log(row.slug + ": ERROR: " + err);
  }
}

function convertPromYmlToSupplementalXml_(sourceXml) {
  var categories = parseCategories_(sourceXml);
  var items = parseOffers_(sourceXml, categories);
  var xml = buildSupplementalFeedXml_(items.writtenItems);
  return {
    xml: xml,
    itemsTotal: items.total,
    itemsWritten: items.writtenItems.length,
    missingCategory: items.missingCategory
  };
}

function parseCategories_(xmlText) {
  var categories = {};
  var re = /<category\b([^>]*)>([\s\S]*?)<\/category>/gi;
  var m;

  while ((m = re.exec(xmlText)) !== null) {
    var attrs = parseAttrs_(m[1]);
    var id = String(attrs.id || "").trim();
    if (!id) continue;
    categories[id] = {
      id: id,
      parentId: String(attrs.parentId || attrs.parentid || "").trim(),
      name: normalizeText_(decodeXml_(extractText_(m[2])))
    };
  }

  return categories;
}

function parseOffers_(xmlText, categories) {
  var out = [];
  var missingCategory = 0;
  var total = 0;
  var re = /<offer\b([^>]*)>[\s\S]*?<\/offer>/gi;
  var m;

  while ((m = re.exec(xmlText)) !== null) {
    total++;
    var block = m[0];
    var attrs = parseAttrs_(m[1]);
    var id = normalizeText_(decodeXml_(attrs.id || extractTag_(block, "id")));
    var categoryId = normalizeText_(decodeXml_(extractTag_(block, "categoryId")));
    if (!id || !categoryId) {
      missingCategory++;
      continue;
    }

    var productType = buildCategoryPath_(categoryId, categories);
    if (!productType) {
      missingCategory++;
      continue;
    }

    out.push({
      id: id,
      productType: productType
    });
  }

  return {
    total: total,
    writtenItems: out,
    missingCategory: missingCategory
  };
}

function buildCategoryPath_(categoryId, categories) {
  var parts = [];
  var seen = {};
  var currentId = String(categoryId || "").trim();

  while (currentId && categories[currentId] && !seen[currentId]) {
    seen[currentId] = true;
    if (categories[currentId].name) parts.unshift(categories[currentId].name);
    currentId = categories[currentId].parentId;
  }

  return normalizeProductType_(parts.join(" > "));
}

function buildSupplementalFeedXml_(items) {
  var lines = [
    '<?xml version="1.0" encoding="UTF-8"?>',
    '<rss version="2.0" xmlns:g="http://base.google.com/ns/1.0">',
    "  <channel>",
    "    <title>Prom product type supplemental feed</title>"
  ];

  for (var i = 0; i < items.length; i++) {
    lines.push("    <item>");
    lines.push("      <g:id>" + escapeXml_(items[i].id) + "</g:id>");
    lines.push("      <g:product_type>" + escapeXml_(items[i].productType) + "</g:product_type>");
    lines.push("    </item>");
  }

  lines.push("  </channel>");
  lines.push("</rss>");
  return lines.join("\n");
}

function fetchText_(url) {
  var response = UrlFetchApp.fetch(url, {
    muteHttpExceptions: true,
    followRedirects: true,
    headers: {
      "Accept-Encoding": "gzip"
    }
  });

  var code = response.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error("Could not fetch source feed. HTTP " + code);
  }

  var blob = response.getBlob();
  var bytes = blob.getBytes();
  if (bytes && bytes.length >= 2 && bytes[0] === 0x1f && bytes[1] === 0x8b) {
    return Utilities.ungzip(blob).getDataAsString("UTF-8");
  }
  return blob.getDataAsString("UTF-8");
}

function upsertOutputFile_(row, xml) {
  var folder = getOrCreateFolder_(OUTPUT_FOLDER_NAME);
  var fileName = row.slug + "-product-type-supplemental.xml";

  if (row.drive_file_id) {
    try {
      var existing = DriveApp.getFileById(row.drive_file_id);
      existing.setContent(xml);
      existing.setName(fileName);
      return existing;
    } catch (err) {
      Logger.log("Existing file could not be updated, creating a new one: " + err);
    }
  }

  return folder.createFile(fileName, xml, MimeType.XML);
}

function buildOutputUrl_(slug, fileId) {
  if (WEB_APP_BASE_URL && /^https?:\/\//i.test(WEB_APP_BASE_URL)) {
    return WEB_APP_BASE_URL.replace(/\?+$/, "") + "?slug=" + encodeURIComponent(slug);
  }
  return "https://drive.google.com/uc?export=download&id=" + encodeURIComponent(fileId);
}

function getConfigRows_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var values = sheet.getRange(2, 1, lastRow - 1, COLS.length).getValues();
  var rows = [];
  for (var i = 0; i < values.length; i++) {
    var obj = { rowNumber: i + 2 };
    for (var j = 0; j < COLS.length; j++) {
      obj[COLS[j]] = values[i][j];
    }
    obj.slug = slugify_(obj.slug || obj.site);
    obj.source_feed_url = String(obj.source_feed_url || "").trim();
    obj.drive_file_id = String(obj.drive_file_id || "").trim();
    rows.push(obj);
  }
  return rows;
}

function validateRow_(row) {
  if (!row.slug) throw new Error("Missing slug.");
  if (!/^https?:\/\//i.test(row.source_feed_url)) {
    throw new Error("source_feed_url must be a direct HTTP/HTTPS URL.");
  }
}

function getControlSheet_() {
  var ss = SpreadsheetApp.openByUrl(CONTROL_SPREADSHEET_URL);
  var sheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(CONFIG_SHEET_NAME);
  return sheet;
}

function ensureHeader_(sheet) {
  var header = sheet.getRange(1, 1, 1, COLS.length).getValues()[0];
  var needsHeader = false;
  for (var i = 0; i < COLS.length; i++) {
    if (header[i] !== COLS[i]) {
      needsHeader = true;
      break;
    }
  }
  if (needsHeader) {
    sheet.getRange(1, 1, 1, COLS.length).setValues([COLS]);
  }
}

function writeStatus_(sheet, rowNumber, status, date, total, written, missing, error) {
  sheet.getRange(rowNumber, colIndex_("last_status")).setValue(status);
  sheet.getRange(rowNumber, colIndex_("last_updated")).setValue(date);
  sheet.getRange(rowNumber, colIndex_("items_total")).setValue(total);
  sheet.getRange(rowNumber, colIndex_("items_written")).setValue(written);
  sheet.getRange(rowNumber, colIndex_("missing_category")).setValue(missing);
  sheet.getRange(rowNumber, colIndex_("error")).setValue(error);
}

function colIndex_(name) {
  for (var i = 0; i < COLS.length; i++) {
    if (COLS[i] === name) return i + 1;
  }
  throw new Error("Unknown column: " + name);
}

function getOrCreateFolder_(name) {
  var folders = DriveApp.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(name);
}

function parseAttrs_(s) {
  var attrs = {};
  var re = /([:\w-]+)\s*=\s*("([^"]*)"|'([^']*)')/g;
  var m;
  while ((m = re.exec(String(s || ""))) !== null) {
    attrs[m[1]] = decodeXml_(m[3] != null ? m[3] : m[4]);
  }
  return attrs;
}

function extractTag_(chunk, tagName) {
  var re = new RegExp("<" + escapeRe_(tagName) + "\\b[^>]*>([\\s\\S]*?)<\\/" + escapeRe_(tagName) + ">", "i");
  var m = re.exec(chunk);
  if (!m) return "";
  var value = m[1] || "";
  var cdata = /<!\[CDATA\[([\s\S]*?)\]\]>/i.exec(value);
  if (cdata) value = cdata[1];
  return extractText_(value);
}

function stripTags_(s) {
  return String(s || "").replace(/<[^>]*>/g, "");
}

function extractText_(s) {
  var value = String(s || "");
  var cdata = /<!\[CDATA\[([\s\S]*?)\]\]>/i.exec(value);
  if (cdata) value = cdata[1];
  return stripTags_(value);
}

function normalizeText_(s) {
  return String(s || "")
    .replace(/\u00A0/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function normalizeProductType_(s) {
  return normalizeText_(s).replace(/\s*>\s*/g, " > ");
}

function decodeXml_(s) {
  return String(s || "")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&")
    .replace(/&quot;/g, "\"")
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, function(_, n) {
      return String.fromCharCode(Number(n));
    });
}

function escapeXml_(s) {
  return String(s || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function escapeRe_(s) {
  return String(s).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function slugify_(s) {
  var slug = String(s || "")
    .toLowerCase()
    .replace(/^https?:\/\//, "")
    .replace(/^www\./, "")
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "");
  return slug || "feed";
}
