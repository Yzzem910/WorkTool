/**
 * generate_doc_from_sheets.gs
 *
 * Apps Script to export a Product Tracker Google Sheet into a nicely formatted Google Doc.
 *
 * Usage:
 * 1) In Google Drive open your Product Tracker spreadsheet (either upload the Excel and open with Google Sheets,
 *    or open the converted Google Sheet that contains the sheets: Products, Listings, Ads, Promotions,
 *    ListingMaintenance, Tasks, Lists, Dashboard).
 * 2) Extensions -> Apps Script. Create a new project (or replace the default Code.gs) and paste this file.
 * 3) Save, then select the function createDocFromProductTracker and Run. Authorize when prompted.
 * 4) After the script completes it will open an alert with a link to the created Google Doc.
 *
 * What it does:
 * - Iterates a predefined set of sheet names (Products, Listings, Ads, Promotions, ListingMaintenance, Tasks, Lists, Dashboard).
 * - For each sheet it copies the header row and up to 'maxRowsPerSheet' data rows into the Google Doc as a table.
 * - For the Dashboard sheet, if present, it copies the used range as text/KPI lines.
 * - The created Google Doc is placed in your Drive root and a URL is returned.
 *
 * Notes & tips:
 * - The script runs in the context of the active spreadsheet; make sure you opened the intended file before running.
 * - You can adjust maxRowsPerSheet to include more/less rows in the Doc export.
 * - If you want full export of very large sheets, consider exporting CSVs instead or increase quotas accordingly.
 */

function createDocFromProductTracker() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    SpreadsheetApp.getUi().alert('未检测到活动的 Google Spreadsheet。请先在 Google Sheets 中打开 Product Tracker，然后再运行此脚本。');
    return;
  }

  var sheetNames = ['Products','Listings','Ads','Promotions','ListingMaintenance','Tasks','Lists','Dashboard'];
  var maxRowsPerSheet = 200; // change if you want more rows exported per sheet

  var now = new Date();
  var title = 'Product Tracker Export - ' + Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  var doc = DocumentApp.create(title);
  var body = doc.getBody();

  // Cover
  body.appendParagraph('Product Tracker Export').setHeading(DocumentApp.ParagraphHeading.TITLE);
  body.appendParagraph('Generated from spreadsheet: ' + ss.getName()).setItalic(true);
  body.appendParagraph('Date: ' + Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'));
  body.appendHorizontalRule();

  sheetNames.forEach(function(name) {
    var sh = ss.getSheetByName(name);
    if (!sh) {
      // skip if sheet absent
      return;
    }

    var dataRange = sh.getDataRange();
    var values = dataRange.getValues();
    if (!values || values.length === 0) return;

    // Add section heading
    body.appendParagraph(name).setHeading(DocumentApp.ParagraphHeading.HEADING1);

    // Prepare table data: header + limited rows
    var header = values[0];
    var rows = [];
    rows.push(header);

    var rowCount = Math.min(values.length - 1, maxRowsPerSheet);
    for (var r = 1; r <= rowCount; r++) {
      rows.push(values[r]);
    }

    // Convert rows to string matrix for doc table: all values to string and trim
    var table = rows.map(function(row) {
      return row.map(function(cell) {
        if (cell === null || cell === undefined) return '';
        // Format dates nicely
        if (cell instanceof Date) return Utilities.formatDate(cell, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        return String(cell);
      });
    });

    // Insert table into doc. Document tables require a rectangular array; ensure consistent column count.
    var maxCols = header.length;
    // Ensure every row length equals maxCols
    table = table.map(function(r) {
      var out = r.slice(0, maxCols);
      while (out.length < maxCols) out.push('');
      return out;
    });

    try {
      body.appendTable(table);
    } catch (e) {
      // If table insertion fails (very wide table), fall back to bullet list of rows
      body.appendParagraph('（表格过宽，以下为文本版数据预览）');
      for (var i = 1; i < table.length; i++) {
        body.appendParagraph(table[i].join(' | '));
      }
    }

    body.appendParagraph('---');
  });

  // Add a small footer with link to sheet
  body.appendParagraph('Source spreadsheet: ' + ss.getUrl()).setLinkUrl(ss.getUrl());

  doc.saveAndClose();

  var url = doc.getUrl();
  SpreadsheetApp.getUi().alert('Google Doc 已创建：' + url);

  // Also log URL in the sheet (Dashboard first cell) if Dashboard exists
  var dash = ss.getSheetByName('Dashboard');
  if (dash) {
    dash.getRange('A1').setValue('Last exported Doc URL:');
    dash.getRange('B1').setValue(url);
  }
}
