TERM SEARCH — OFFLINE

1. Keep this entire folder together (including the vendor folder).
2. Open index.html in Chrome, Edge, or another modern browser.
3. Drag one or more Excel workbooks onto the import area, or click it to choose files.
4. Paste or type Chinese into the search box. No Enter key is needed.
5. Click the circled ? beside the introduction for the functions and keyboard
   shortcuts guide. Close it with the × button or Escape.

Multiple workbooks
------------------
Imported workbooks form a local library. Use **Manage files** to remove an
unwanted workbook without affecting the others. Drag the nine-dot handle to
set its priority, or focus the handle and press Up/Down. The Search on/off
button immediately to the left of Remove excludes a file from searches without
removing it. File order and on/off settings are remembered locally.
Circled workbook numbers have their own boxes in the library.
When multiple workbooks are loaded, matching numbers appear beside sheet names
in results. With one workbook, results show only the sheet name.
Numbers follow the library order.

The default Smart search checks every enabled workbook and is the most complete
choice. For very large libraries, choose **Fast — stop at the first file with
an exact match**. It checks exact matches in priority order and stops at the
first matching workbook; if there is no exact match, it searches all files for
close matches. This keeps normal searches responsive without silently hiding
results by default.

The arrows beside the Sheet heading cycle through All files, then each enabled
workbook in library order, then back to All files. The label below Sheet shows
the current filter; hover over it to read the full filename. Both Smart and Fast
respect this filter before ranking results and applying the 100-row display cap.
Disabling or removing the selected workbook resets the filter to All files.

Pinyin search
-------------
Click 拼音 at the top right of the search card to search only Chinese text in
the Native column by pronunciation. English text, Translation and aliases do
not produce pinyin hits. Turn it off to restore normal search with the same query.
Examples: zichan, zi chan, zīchǎn and zi1chan3 find 资产 (also 資產).
Loose z/zh, c/ch and s/sh spellings are accepted, so zican also finds 资产.
Partial spellings match while typing, starting at a syllable boundary. Separate
alternatives with commas; spaces join syllables. Use ü, v or u: for ü.
Matching Native characters are highlighted; complete standard spellings rank
first. Pronunciation uses the bundled pinyin-pro 3.29.4 dictionary and works
offline. Uncommon names or context-dependent pronunciations may differ.
Pinyin mode and the Sheet filter reset when the page is reopened.

Privacy
-------
The page has no upload, server, or Internet request. The workbook is read only
in your browser and the original Excel file is never modified. The imported
library and its settings are saved in this browser's local IndexedDB storage.

Workbook format
---------------
The page checks every sheet. It finds the
first header row containing:
  - Native
  - Translation, or Approved Translation

It ignores other columns. Blank Native/Translation rows are skipped.

Using the results
-----------------
* The 全选 pill inside the search bar toggles select-all on a single click.
  Green enables it; red restores normal clicking. Double-click still selects
  a word normally, and dragging still selects part of the text. It starts off.
* Pasting a search, or typing into an empty search box, returns the table to
  the top. Continuing to type, deleting, clearing, and undo/redo keep the
  current offset where results allow. Undo after clearing restores the offset.
* Native partial-match ranking counts repeated matches only once, including
  Traditional/Simplified equivalents and Simplified aliases. Chinese/mixed
  searches count distinct matching characters, including digits.
  For example, searching 资金项目 gives 资金其他内容资金更多文字资金 a count of 2.
  Searching 资金123456 gives 资金1234资金4567资金 a count of 8 (资, 金, and 1–6).
  English/numeric searches count each matched word/number once, weighted by
  its length. All occurrences remain highlighted; exact matches stay first.
* Single-click a Native or Translation value to copy the complete value.
* Press Down in the search box to select the first result (Up selects the last).
  Use Up/Down to move between highlighted rows and Enter to copy the selected
  translation. Escape, or Up on the first row, returns to the search box.
  The table scrolls to keep the selected row visible. These shortcuts do not
  interfere with Chinese composition or editing a cell.
* Select a row with the keyboard or click it to show its matching count next
  to the total result count. "Selected Native: 4 distinct matches" describes
  only that row, not all results. Exact Native/Translation matches are labeled.
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

Developer checks
----------------
With Node.js installed, run: node --test tests/search.test.cjs
These checks use a simulated DOM; they do not replace a browser interaction test.

Pinyin dependency
-----------------
pinyin-pro 3.29.4 (MIT), https://github.com/zh-lx/pinyin-pro
Bundled browser build: vendor/pinyin-pro.js
License: vendor/pinyin-pro.LICENSE
Source package: https://registry.npmjs.org/pinyin-pro/-/pinyin-pro-3.29.4.tgz
Package SHA-512 (base64): SPXpDT2cHEy+d26V1RXYMlVzXN42hotFAak1fzyWPi4o2dKXb61UqD4pzxDJHwk6gbv8vQ6EfErd+hYX0Qhzug==
