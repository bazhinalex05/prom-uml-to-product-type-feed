# Prom UML/YML -> Merchant Center product_type supplemental feed

Google Apps Script converter for Prom.ua UML/YML feeds.

## What It Does

The script reads a Prom UML/YML feed, builds the website category path from:

- `<category id="..." parentId="...">Category name</category>`
- `<offer id="..."><categoryId>...</categoryId></offer>`

Then it creates a Merchant Center supplemental XML feed:

```xml
<item>
  <g:id>PRODUCT_ID</g:id>
  <g:product_type>Parent category > Child category</g:product_type>
</item>
```

Merchant Center matches this supplemental feed to the main feed by `g:id`.

## Control Sheet Columns

Run `setupControlSheet()` once. It creates a `Feeds` sheet with:

| column | meaning |
|---|---|
| `enabled` | `TRUE` rows are processed |
| `site` | human readable site name/domain |
| `slug` | stable short feed key, used in output file name and Web App URL |
| `source_feed_url` | direct Prom UML/YML feed URL |
| `output_feed_url` | generated Merchant Center supplemental feed URL |
| `drive_file_id` | internal Drive file ID, reused on next updates |
| `last_status` | `OK`, `ERROR`, or `RUNNING` |
| `last_updated` | last processing time |
| `items_total` | total `<offer>` blocks read |
| `items_written` | items written to supplemental feed |
| `missing_category` | offers skipped because ID/category path was missing |
| `error` | last error text |

## Setup

1. Create a Google Sheet for the feed registry.
2. Open `Extensions -> Apps Script`.
3. Paste `script.js`.
4. Set `CONTROL_SPREADSHEET_URL` to the registry sheet URL.
5. Run `setupControlSheet()`.
6. Fill rows in the `Feeds` sheet.
7. Run `runAllFeeds()`.
8. Put `output_feed_url` into Merchant Center as a supplemental feed URL.

## Daily Updates

In Apps Script, add a time-driven trigger for `runAllFeeds()`.

Recommended first test:

- run once manually;
- check `items_total`, `items_written`, and `missing_category`;
- open `output_feed_url`;
- connect the URL as a Merchant Center supplemental feed.

## Output URL Options

By default, the script writes XML files to Google Drive and outputs a direct Drive download URL.

For a cleaner URL:

1. Deploy the script as a Web App.
2. Copy the Web App URL.
3. Set `WEB_APP_BASE_URL` in the script.
4. Run `runAllFeeds()` again.

Then `output_feed_url` will look like:

```text
WEB_APP_URL?slug=site-slug
```

## Limits

This is designed for small and medium Prom feeds that fit Google Apps Script limits. If a feed is too large or processing times out, keep the same control sheet but move the converter runtime to GitHub Actions or Cloudflare.
