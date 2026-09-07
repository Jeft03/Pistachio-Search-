TERM SEARCH — OFFLINE

1. Keep this entire folder together (including the vendor folder).
2. Open index.html in Chrome, Edge, or another modern browser.
3. Drag your Excel workbook onto the import area, or click it to choose the file.
4. Paste or type Chinese into the search box. No Enter key is needed.

Privacy
-------
The page has no upload, server, or Internet request. The workbook is read only
in your browser and is never modified. It is not saved by the page.

Workbook format
---------------
The page uses every sheet except the final three workbook sheets. It finds the
first header row containing:
  - Native
  - Translation, or Approved Translation

It ignores other columns. Blank Native/Translation rows are skipped.

Using the results
-----------------
* Single-click a Native or Translation value to copy the complete value.
* Double-click it to temporarily select or edit text. Click elsewhere or press
  Escape to leave that view; a new search restores the original imported value.
* Search terms separated by spaces, commas, etc. use OR matching. For Chinese
  text pasted together (e.g. 资产负债), the page also recognizes known Native
  terms occurring inside it (e.g. 资产 and 负债).

Automatic translation
---------------------
This offline version deliberately does not generate new translations. A genuine
automatic fallback needs either an online translation provider or a separately
installed local language model; neither is bundled, and terminology is never
sent from this page.
