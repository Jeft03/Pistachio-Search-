TERM SEARCH — OFFLINE

1. Keep this entire folder together (including the vendor folder).
2. Open index.html in Chrome, Edge, or another modern browser.
3. Drag one or more Excel workbooks onto the import area, or click it to choose files.
4. Paste or type Chinese into the search box. No Enter key is needed.

Multiple workbooks
------------------
Imported workbooks form a local library. Use **Manage files** to remove an
unwanted workbook without affecting the others, or move a workbook up/down to
set its priority. Circled workbook numbers have their own boxes in the library.
When multiple workbooks are loaded, matching numbers appear beside sheet names
in results. With one workbook, results show only the sheet name.
Numbers follow the library order.

The default Smart search checks every kept workbook and is the most complete
choice. For very large libraries, choose **Fast — stop at the first file with
an exact match**. It checks exact matches in priority order and stops at the
first matching workbook; if there is no exact match, it searches all files for
close matches. This keeps normal searches responsive without silently hiding
results by default.

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
* Drag directly across Native or Translation text to select any part of it.
  Selection follows the cursor, including partial words. Release the mouse
  to copy automatically. Dragging works in either direction.
* Right-click a value and choose Edit to change it.
* Search terms separated by spaces, commas, etc. use OR matching. For Chinese
  text pasted together (e.g. 资产负债), the page also recognizes known Native
  terms occurring inside it (e.g. 资产 and 负债).

Automatic translation
---------------------
This offline version deliberately does not generate new translations. A genuine
automatic fallback needs either an online translation provider or a separately
installed local language model; neither is bundled, and terminology is never
sent from this page.
