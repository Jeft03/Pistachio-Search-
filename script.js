/* Runs only in the browser. No fetch, server, or workbook write is used. */
const $ = (id) => document.getElementById(id);
const dropZone = $('dropZone'), fileInput = $('fileInput'), searchInput = $('searchInput');
let terms = [], workbooks = [], exactIndexes = [], searchTimer = null, activeCell = null;
const openCCReady = Boolean(window.OpenCC && OpenCC.Converter);
const convertTraditionalToSimplified = openCCReady ? OpenCC.Converter({from:'tw', to:'cn'}) : value => value;
const convertSimplifiedToTraditional = openCCReady ? OpenCC.Converter({from:'cn', to:'tw'}) : value => value;

dropZone.addEventListener('click', () => fileInput.click());
dropZone.addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); fileInput.click(); } });
fileInput.addEventListener('change', e => { if (e.target.files.length) importWorkbooks([...e.target.files]); e.target.value = ''; });
['dragenter','dragover'].forEach(type => dropZone.addEventListener(type, e => { e.preventDefault(); dropZone.classList.add('drag'); }));
['dragleave','drop'].forEach(type => dropZone.addEventListener(type, e => { e.preventDefault(); dropZone.classList.remove('drag'); }));
dropZone.addEventListener('drop', e => { const files = [...e.dataTransfer.files].filter(f => /\.(xlsx|xlsm|xlsb)$/i.test(f.name)); if (files.length) importWorkbooks(files); else setStatus('Please choose Excel .xlsx, .xlsm, or .xlsb files.', 'error'); });
searchInput.addEventListener('input', scheduleRenderResults);
$('toggleLibrary').addEventListener('click', () => { const panel = $('libraryPanel'), button = $('toggleLibrary'); panel.hidden = !panel.hidden; button.setAttribute('aria-expanded', String(!panel.hidden)); button.textContent = panel.hidden ? 'Manage files' : 'Close manager'; });
$('searchScope').addEventListener('change', renderResults);
setUpSheetColumnResize();

async function importWorkbooks(files) {
  let imported = 0;
  for (const file of files) {
    try { if (await addWorkbook(file.name, await file.arrayBuffer())) imported++; }
    catch (error) { console.error(error); setStatus(`Could not read ${file.name}: ${error.message}`, 'error'); }
  }
  if (imported) { searchInput.focus(); renderResults(); }
}
async function addWorkbook(name, buffer, {persist = true} = {}) {
  if (!window.JSZip) return setStatus('The local Excel reader is missing. Keep the vendor folder beside this page.', 'error');
  try {
    setStatus(`Reading ${name} locally…`);
    const zip = await JSZip.loadAsync(buffer);
    const parsed = await readXlsx(zip);
    const bookTerms = prepareSearchTerms(parsed.terms.map(term => ({...term, workbookId:name, workbookName:name})));
    if (!bookTerms.length) throw new Error('No sheet with both a Native column and a Translation/Approved Translation column was found.');
    const previous = workbooks.find(item => item.id === name);
    const book = {id:name, name, buffer, terms:bookTerms, usedSheets:parsed.usedSheets, priority:previous ? previous.priority : workbooks.length, saved:Date.now()};
    const existing = workbooks.findIndex(item => item.id === book.id);
    if (existing >= 0) workbooks.splice(existing, 1, book); else workbooks.push(book);
    rebuildSearchData();
    $('searchArea').hidden = $('resultsSection').hidden = false;
    $('fallback').hidden = true;
    if (persist) await saveBook(book);
    renderWorkbookLibrary();
    setStatus(`${name}: ${bookTerms.length.toLocaleString()} terms read from ${parsed.usedSheets} sheet${parsed.usedSheets === 1 ? '' : 's'}. ${workbooks.length.toLocaleString()} workbook${workbooks.length === 1 ? '' : 's'} are ready to search.${openCCReady ? ' Traditional/Simplified matching is ready.' : ' Warning: Traditional/Simplified converter did not load; keep the vendor folder with this page.'}`, openCCReady ? 'good' : 'error');
    return true;
  } catch (error) { console.error(error); setStatus(`Could not read ${name}: ${error.message}`, 'error'); return false; }
}

async function readXlsx(zip) {
  const xml = async path => new DOMParser().parseFromString(await zip.file(path).async('text'), 'application/xml');
  const workbook = await xml('xl/workbook.xml');
  const rels = await xml('xl/_rels/workbook.xml.rels');
  const relation = Object.fromEntries([...rels.querySelectorAll('Relationship')].map(r => [r.getAttribute('Id'), r.getAttribute('Target')]));
  const shared = zip.file('xl/sharedStrings.xml') ? [...(await xml('xl/sharedStrings.xml')).querySelectorAll('si')].map(readRichText) : [];
  const sheets = [...workbook.querySelectorAll('sheets > sheet')];
  const termsOut = []; let usedSheets = 0;
  for (const sheet of sheets) {
    const target = relation[sheet.getAttribute('r:id')]; if (!target) continue;
    const path = target.startsWith('/') ? target.slice(1) : `xl/${target.replace(/^\.\//, '')}`;
    const rows = parseRows(await xml(path), shared);
    const headerIndex = rows.findIndex(row => row.some(v => /^native$/i.test(v.trim())) && row.some(v => /^(approved )?translation$/i.test(v.trim())));
    if (headerIndex < 0) continue;
    const headers = rows[headerIndex].map(v => v.trim().toLowerCase());
    const nativeIndex = headers.indexOf('native'); const translationIndex = headers.findIndex(v => /^(approved )?translation$/.test(v)); const simplifiedIndex = headers.indexOf('simplified');
    let added = 0;
    rows.slice(headerIndex + 1).forEach((row, index) => { const native = (row[nativeIndex] || '').trim(), translation = (row[translationIndex] || '').trim(), simplified = simplifiedIndex >= 0 ? (row[simplifiedIndex] || '').trim() : ''; if (native && translation) { termsOut.push({sheet: sheet.getAttribute('name'), native, translation, aliases:simplified && simplified !== native ? [simplified] : [], originalNative:native, originalTranslation:translation, order: termsOut.length, row: headerIndex + index + 2, sheetPath:path, nativeIndex, translationIndex}); added++; } });
    if (added) usedSheets++;
  }
  return {terms: termsOut, usedSheets, sheets:[...new Set(termsOut.map(t => t.sheet))]};
}
function readRichText(node) { return [...node.querySelectorAll('t')].map(t => t.textContent).join(''); }
function parseRows(doc, shared) {
  return [...doc.querySelectorAll('sheetData > row')].map(row => {
    const values = []; [...row.querySelectorAll('c')].forEach(cell => {
      const ref = cell.getAttribute('r') || ''; const index = columnNumber(ref.replace(/\d/g, ''));
      const type = cell.getAttribute('t'); let value = '';
      if (type === 'inlineStr') value = readRichText(cell); else { const v = cell.querySelector('v'); value = v ? v.textContent : ''; if (type === 's') value = shared[Number(value)] ?? ''; }
      values[index] = String(value);
    }); return values;
  });
}
function columnNumber(letters) { let n = 0; for (const c of letters) n = n * 26 + c.charCodeAt(0) - 64; return n - 1; }

function scheduleRenderResults() {
  clearTimeout(searchTimer);
  searchTimer = setTimeout(() => { searchTimer = null; renderResults(); }, 120);
}
function prepareSearchTerms(items) { return items.map(prepareSearchTerm); }
function rebuildSearchData() {
  workbooks.forEach((book, index) => { book.priority = index; });
  terms = workbooks.flatMap(book => book.terms);
  terms.forEach((term, index) => { term.order = index; });
  exactIndexes = workbooks.map(book => {
    const index = new Map();
    for (const term of book.terms) {
      for (const value of [term.native, term.translation, ...(term.aliases || [])]) {
        const key = canonicalTerm(value);
        if (!key) continue;
        const matches = index.get(key) || [];
        matches.push(term);
        index.set(key, matches);
      }
    }
    return index;
  });
}
function prepareSearchTerm(term) {
  const aliasText = (term.aliases || []).join(' ');
  term._search = {
    normalised: [normalise(term.native), normalise(term.translation), normalise(aliasText)],
    canonical: [canonicalTerm(term.native), canonicalTerm(term.translation), canonicalTerm(aliasText)],
    raw: [term.native, term.translation, aliasText]
  };
  return term;
}
function makeSearchTokenData(tokens) {
  return tokens.map(value => ({value, normalised:normalise(value), canonical:canonicalTerm(value), chinese:hasChinese(value)}));
}
function hasChineseSubsequence(text, token) {
  const chars = [...token.replace(/\s/g, '')]; let cursor = 0;
  for (const char of chars) { const at = text.indexOf(char, cursor); if (at < 0) return false; cursor = at + 1; }
  return chars.length > 0;
}
function isCandidate(term, tokenData) {
  const search = term._search || prepareSearchTerm(term)._search;
  return tokenData.some(token => search.canonical.some(text => text.includes(token.canonical)) || (token.chinese && search.raw.some(text => hasChineseSubsequence(text, token.value))));
}
function renderResults() {
  const query = searchInput.value.trim(), body = $('resultsBody'); body.textContent = '';
  const tokens = makeSearchTokens(query);
  const tokenData = makeSearchTokenData(tokens), queryCanonical = canonicalTerm(query), normalisedQuery = normalise(query);
  const displayed = [], maximumResults = 100;
  let candidates = [];
  if (query) {
    const firstExactIndex = $('searchScope').value === 'first-exact' ? exactIndexes.find(index => index.has(queryCanonical)) : null;
    // Fast mode never walks later workbooks after an exact hit. All other searches use the full local index.
    candidates = firstExactIndex ? firstExactIndex.get(queryCanonical) : terms.filter(term => isCandidate(term, tokenData));
  }
  const matchingCount = candidates.length;
  for (const t of candidates) {
    const search = t._search || prepareSearchTerm(t)._search;
    const nativeMatch = matchText(t.native, tokenData, search.normalised[0]), translationMatch = matchText(t.translation, tokenData, search.normalised[1]), aliasMatch = matchText((t.aliases || []).join(' '), tokenData, search.normalised[2]);
    const candidate = {...t, nativeMatch, translationMatch, aliasMatch, bestLength: Math.max(nativeMatch.bestLength, translationMatch.bestLength, aliasMatch.bestLength)};
    candidate.rankScore = score(candidate, queryCanonical, normalisedQuery);
    if (displayed.length < maximumResults) displayed.push(candidate);
    else { let weakest = 0; for (let i = 1; i < displayed.length; i++) if (compareResults(displayed[i], displayed[weakest]) > 0) weakest = i; if (compareResults(candidate, displayed[weakest]) < 0) displayed[weakest] = candidate; }
  }
  displayed.sort(compareResults);
  $('resultCount').textContent = query ? `${matchingCount.toLocaleString()} matching term${matchingCount === 1 ? '' : 's'}` : 'Type or paste Chinese or English to start searching';
  $('emptyState').hidden = !query || matchingCount > 0; $('fallback').hidden = !query || matchingCount > 0;
  for (const term of displayed) {
    const exactNative = isExactNativeMatch(term.native, query);
    const tr = document.createElement('tr');
    const source = document.createElement('td');
    if (workbooks.length > 1) source.append(workbookBadge(workbooks.findIndex(book => book.id === term.workbookId) + 1), ' · ');
    source.append(term.sheet);
    tr.append(source, copyCell(term, 'native', term.nativeMatch.ranges, exactNative), copyCell(term, 'translation', term.translationMatch.ranges));
    body.append(tr);
  }
}
function compareResults(a, b) { return b.rankScore - a.rankScore || b.bestLength - a.bestLength || a.order - b.order; }
function makeSearchTokens(query) {
  const direct = query.trim().split(/[\s,，;；、|/]+/).filter(Boolean);
  const symbolTokens = [...query].filter(char => Object.hasOwn(punctuationEquivalents, char) || [':',';','\'', '"', ',', '.', '(', ')', '-'].includes(char));
  const joinedChinese = direct.every(hasChinese) && direct.length > 1 ? [direct.join('')] : [];
  const pieces = [];
  for (const word of [...direct, ...joinedChinese]) {
    if (!hasChinese(word)) continue;
    const chars = [...word.replace(/\s/g, '')];
    for (let size = Math.min(chars.length, 12); size >= 2; size--) for (let start = 0; start + size <= chars.length; start++) pieces.push(chars.slice(start, start + size).join(''));
  }
  const variants = [...direct, ...joinedChinese, ...pieces, ...symbolTokens].flatMap(chineseVariants);
  return [...new Set(variants)].sort((a,b) => b.length - a.length);
}
function chineseVariants(value) { if (!hasChinese(value)) return [value]; return [value, convertTraditionalToSimplified(value), convertSimplifiedToTraditional(value)]; }
function matchText(text, tokens, normalisedText = normalise(text)) {
  const ranges = []; let bestLength = 0; let matched = false; let wholeQuery = false;
  for (const token of tokens) {
    const contiguous = findContiguousRanges(text, token.value, normalisedText, token.normalised);
    const fuzzy = !contiguous.length && token.chinese ? findChineseSubsequenceRanges(text, token.value) : [];
    const current = contiguous.length ? contiguous : fuzzy;
    if (current.length) { matched = true; bestLength = Math.max(bestLength, token.value.length); ranges.push(...current); }
  }
  const merged = mergeRanges(ranges);
  return {matched, ranges: merged, bestLength, coverage: merged.reduce((total, range) => total + range[1] - range[0], 0), wholeQuery};
}
function score(term, queryCanonical, normalisedQuery) {
  const search = term._search || prepareSearchTerm(term)._search;
  const exactNative = search.canonical[0] === queryCanonical, exactTranslation = search.canonical[1] === queryCanonical;
  if (exactNative) return 1000000;
  if (exactTranslation) return 950000;
  const nativeWhole = search.canonical[0].includes(queryCanonical), translationWhole = search.canonical[1].includes(queryCanonical);
  if (nativeWhole) return 500000 + normalisedQuery.length;
  if (translationWhole) return 450000 + normalisedQuery.length;
  const fields = [{text:term.native,match:term.nativeMatch},{text:term.translation,match:term.translationMatch},{text:(term.aliases || []).join(' '),match:term.aliasMatch}];
  const quality = Math.max(...fields.map(({text,match}) => match.coverage * 10000 - Math.max(0, [...text].length - match.coverage) * 300));
  return quality + term.bestLength * 10;
}
function cell(text) { const td = document.createElement('td'); td.textContent = text; return td; }
function workbookBadge(number) {
  const badge = document.createElement('span');
  badge.className = 'workbook-number';
  badge.textContent = String(number);
  badge.setAttribute('aria-label', `Workbook ${number}`);
  return badge;
}
function copyCell(term, field, ranges, exactNative = false) {
  const text = term[field];
  const td = document.createElement('td'), div = document.createElement('div'); div.className = 'copy-cell'; div.dataset.original = text;
  const label = document.createElement('span');
  label.innerHTML = highlightRanges(text, ranges);
  div.append(label);
  if (exactNative) { const tick = document.createElement('img'); tick.className = 'exact-tick'; tick.src = 'assets/exact-match-tick.png'; tick.alt = 'Exact match'; tick.title = 'Exact match'; div.append(' ', tick); }
  setUpDragCopy(div, label, text);
  div.addEventListener('contextmenu', event => { event.preventDefault(); activeCell = {term, field, div}; showCellMenu(event.clientX, event.clientY); });
  td.append(div); return td;
}
const punctuationEquivalents = {'：':':','；':';','‘':"'",'’':"'",'“':'"','”':'"','，':',','。':'.','（':'(','）':')','－':'-','—':'-','–':'-','、':','};
function normalise(value) { return [...String(value).toLocaleLowerCase()].map(char => punctuationEquivalents[char] || char).join('').replace(/\s+/g, ' ').trim(); }
function canonicalTerm(value) { return normalise(convertTraditionalToSimplified(value)); }
function isExactNativeMatch(native, query) { return canonicalTerm(native) === canonicalTerm(query); }
function containsNormalised(text, query) { return normalise(text).includes(normalise(query)); }
function hasChinese(value) { return /[\u3400-\u9fff]/.test(value); }
function findContiguousRanges(text, token, lowerText = normalise(text), lowerToken = normalise(token)) {
  const ranges = []; let from = 0, index;
  while ((index = lowerText.indexOf(lowerToken, from)) !== -1) { ranges.push([index, index + token.length]); from = index + token.length; }
  return ranges;
}
function findChineseSubsequenceRanges(text, token) {
  const chars = [...token.replace(/\s/g, '')]; if (!chars.length) return [];
  const points = []; let cursor = 0;
  for (const char of chars) { const at = text.indexOf(char, cursor); if (at < 0) return []; points.push(at); cursor = at + 1; }
  const ranges = []; let start = points[0], end = start + 1;
  points.slice(1).forEach(point => { if (point === end) end++; else { ranges.push([start, end]); start = point; end = point + 1; } }); ranges.push([start, end]); return ranges;
}
function mergeRanges(ranges) {
  const sorted = ranges.slice().sort((a,b) => a[0] - b[0] || b[1] - a[1]), merged = [];
  for (const range of sorted) { const last = merged[merged.length - 1]; if (last && range[0] <= last[1]) last[1] = Math.max(last[1], range[1]); else merged.push([...range]); }
  return merged;
}
function highlightRanges(text, ranges) {
  let out = '', at = 0; for (const [start, end] of ranges) { out += escapeHtml(text.slice(at, start)) + `<mark class="match">${escapeHtml(text.slice(start, end))}</mark>`; at = end; } return out + escapeHtml(text.slice(at));
}
function textPoint(label, offset) {
  const walker = document.createTreeWalker(label, NodeFilter.SHOW_TEXT);
  let node, last;
  while ((node = walker.nextNode())) {
    if (offset <= node.length) return [node, offset];
    offset -= node.length; last = node;
  }
  return last ? [last, last.length] : [label, 0];
}
function offsetAtPointer(label, event) {
  const box = label.getBoundingClientRect();
  const x = Math.max(box.left, Math.min(event.clientX, box.right-1));
  const y = Math.max(box.top, Math.min(event.clientY, box.bottom-1));
  const caret = document.caretPositionFromPoint?.(x,y);
  const range = !caret && document.caretRangeFromPoint?.(x,y);
  const node = caret?.offsetNode || range?.startContainer;
  const offset = caret?.offset ?? range?.startOffset;
  if (!node || !label.contains(node)) return null;
  const prefix = document.createRange();
  prefix.selectNodeContents(label); prefix.setEnd(node, offset);
  return prefix.toString().length;
}
function setUpDragCopy(div, label, text) {
  let suppressClick = false;
  div.addEventListener('click', () => { if (!suppressClick && !div.isContentEditable) copyText(text); });
  div.addEventListener('dblclick', e => { if (!div.isContentEditable) e.preventDefault(); });
  div.addEventListener('dragstart', e => { if (!div.isContentEditable) e.preventDefault(); });
  div.addEventListener('pointerdown', event => {
    if (event.button !== 0 || !event.isPrimary || div.isContentEditable || event.pointerType === 'touch') return;
    suppressClick = false;
    const offset = offsetAtPointer(label, event);
    if (offset === null) return;
    const anchor = offset;
    const id = event.pointerId, x = event.clientX, y = event.clientY;
    let dragging = false, selected = '';
    event.preventDefault(); getSelection()?.removeAllRanges(); div.setPointerCapture(id);
    const move = current => {
      if (current.pointerId !== id || !div.isConnected) return;
      if (!dragging && Math.hypot(current.clientX-x, current.clientY-y) < 4) return;
      dragging = true; suppressClick = true;
      const offset = offsetAtPointer(label, current);
      if (offset === null) return;
      const a = Math.min(anchor, offset), b = Math.max(anchor, offset);
      const range = document.createRange();
      range.setStart(...textPoint(label,a)); range.setEnd(...textPoint(label,b));
      const selection = getSelection(); selection.removeAllRanges(); selection.addRange(range);
      selected = text.slice(a,b);
    };
    const stop = current => {
      if (current.pointerId !== id) return;
      if (current.type === 'pointerup') move(current);
      div.removeEventListener('pointermove',move);
      for (const type of ['pointerup','pointercancel','lostpointercapture']) div.removeEventListener(type,stop);
      if (div.hasPointerCapture(id)) div.releasePointerCapture(id);
      if (current.type === 'pointerup' && div.isConnected && selected.trim()) copyText(selected,'Selected text copied');
      if (current.type !== 'pointerup') suppressClick = true;
    };
    div.addEventListener('pointermove',move);
    for (const type of ['pointerup','pointercancel','lostpointercapture']) div.addEventListener(type,stop);
  });
}
async function copyText(text, message = 'Copied') { try { await navigator.clipboard.writeText(text); showToast(message); } catch { showToast('Select the text and copy manually'); } }
function escapeHtml(s) { return s.replace(/[&<>'"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[c])); }
function setStatus(message, kind='') { const status = $('importStatus'); status.textContent = message; status.className = `status ${kind}`; }
function showToast(message) { const toast = $('toast'); toast.textContent = message; toast.classList.add('show'); setTimeout(() => toast.classList.remove('show'), 1400); }
function setUpSheetColumnResize() {
  const handle = $('sheetResize'), table = $('resultsTable'); let startX = 0, startWidth = 0;
  handle.addEventListener('pointerdown', event => { startX = event.clientX; startWidth = table.querySelector('.sheet-col').getBoundingClientRect().width; handle.setPointerCapture(event.pointerId); handle.classList.add('dragging'); });
  handle.addEventListener('pointermove', event => { if (!handle.hasPointerCapture(event.pointerId)) return; table.style.setProperty('--sheet-width', `${Math.max(65, Math.min(310, startWidth + event.clientX - startX))}px`); });
  const stop = event => { if (handle.hasPointerCapture(event.pointerId)) handle.releasePointerCapture(event.pointerId); handle.classList.remove('dragging'); };
  handle.addEventListener('pointerup', stop); handle.addEventListener('pointercancel', stop);
}

/* Local workbook library. Nothing is sent to a server. */
const DB_NAME = 'pistachio-search-library';
function db() { return new Promise((resolve, reject) => { const r = indexedDB.open(DB_NAME, 2); r.onupgradeneeded = () => { if (!r.result.objectStoreNames.contains('books')) r.result.createObjectStore('books', {keyPath:'id'}); }; r.onsuccess = () => resolve(r.result); r.onerror = () => reject(r.error); }); }
async function saveBook(book) {
  const d = await db(), tx = d.transaction('books', 'readwrite');
  tx.objectStore('books').put({...book, saved:Date.now()});
  await new Promise((resolve, reject) => { tx.oncomplete = resolve; tx.onerror = () => reject(tx.error); });
}
async function deleteBook(id) {
  const d = await db(), tx = d.transaction('books', 'readwrite');
  tx.objectStore('books').delete(id);
  await new Promise((resolve, reject) => { tx.oncomplete = resolve; tx.onerror = () => reject(tx.error); });
}
async function savedBooks() {
  const d = await db(), tx = d.transaction('books', 'readonly'), request = tx.objectStore('books').getAll();
  return new Promise((resolve, reject) => { request.onsuccess = () => resolve(request.result); request.onerror = () => reject(request.error); });
}
async function restoreLibrary() {
  try {
    const books = await savedBooks();
    workbooks = books.map(book => ({...book, id:book.id || book.name, terms:prepareSearchTerms((book.terms || []).map(term => ({...term, workbookId:book.id || book.name, workbookName:book.name}))) })).sort((a, b) => (a.priority ?? Number.MAX_SAFE_INTEGER) - (b.priority ?? Number.MAX_SAFE_INTEGER) || (a.saved || 0) - (b.saved || 0));
    if (!workbooks.length) return;
    rebuildSearchData();
    $('searchArea').hidden = $('resultsSection').hidden = false;
    renderWorkbookLibrary();
    setStatus(`${workbooks.length.toLocaleString()} saved workbook${workbooks.length === 1 ? '' : 's'} restored locally.`, 'good');
  } catch (error) { console.warn('Could not restore local workbook library', error); }
}
function renderWorkbookLibrary() {
  const holder = $('workbookList'); holder.textContent = '';
  $('libraryControls').hidden = !workbooks.length;
  $('librarySummary').textContent = `${workbooks.length} file${workbooks.length === 1 ? '' : 's'} · ${terms.length.toLocaleString()} terms`;
  workbooks.forEach((book, index) => {
    const row = document.createElement('div'); row.className = 'workbook-row';
    const detail = document.createElement('div'); detail.className = 'workbook-detail';
    const numberBox = document.createElement('div'); numberBox.className = 'workbook-number-box';
    numberBox.append(workbookBadge(index + 1));
    const name = document.createElement('strong'); name.textContent = book.name;
    const meta = document.createElement('span'); meta.textContent = `${book.terms.length.toLocaleString()} terms · ${book.usedSheets || new Set(book.terms.map(term => term.sheet)).size} sheet${(book.usedSheets || new Set(book.terms.map(term => term.sheet)).size) === 1 ? '' : 's'}`;
    detail.append(name, meta);
    const actions = document.createElement('div'); actions.className = 'workbook-actions';
    actions.append(workbookButton('↑', 'Move up', () => moveWorkbook(index, -1), index === 0), workbookButton('↓', 'Move down', () => moveWorkbook(index, 1), index === workbooks.length - 1), workbookButton('Remove', `Remove ${book.name}`, () => removeWorkbook(book.id), false, 'remove'));
    row.append(numberBox, detail, actions); holder.append(row);
  });
}
function workbookButton(label, title, action, disabled, className = '') { const button = document.createElement('button'); button.type = 'button'; button.textContent = label; button.title = title; button.setAttribute('aria-label', title); button.disabled = disabled; button.className = className; button.addEventListener('click', action); return button; }
async function moveWorkbook(index, direction) {
  const target = index + direction; if (target < 0 || target >= workbooks.length) return;
  [workbooks[index], workbooks[target]] = [workbooks[target], workbooks[index]];
  rebuildSearchData(); renderWorkbookLibrary(); renderResults();
  await Promise.all(workbooks.map(saveBook));
}
async function removeWorkbook(id) {
  const book = workbooks.find(item => item.id === id); if (!book) return;
  workbooks = workbooks.filter(item => item.id !== id); rebuildSearchData();
  await deleteBook(id); renderWorkbookLibrary();
  if (!workbooks.length) { $('searchArea').hidden = $('resultsSection').hidden = true; setStatus('All workbooks were removed. Choose one or more Excel files to begin.', ''); return; }
  setStatus(`${book.name} was removed. ${workbooks.length} workbook${workbooks.length === 1 ? '' : 's'} remain ready to search.`, 'good'); renderResults();
}
function showCellMenu(x,y){const m=$('cellMenu');m.hidden=false;m.style.left=`${x}px`;m.style.top=`${y}px`;}
document.addEventListener('click',event=>{if(!event.target.closest('#cellMenu'))$('cellMenu').hidden=true;});
$('cellMenu').onclick=async event=>{const action=event.target.dataset.action;if(!action||!activeCell)return;const {term,field,div}=activeCell;const saveTermBook=async()=>{const book=workbooks.find(item=>item.id===term.workbookId);if(book)await saveBook(book);};if(action==='edit'){div.contentEditable='true';div.classList.add('selecting');div.focus();const done=async()=>{term[field]=div.textContent.trim()||term[field];prepareSearchTerm(term);rebuildSearchData();div.contentEditable='false';div.classList.remove('selecting');await saveTermBook();renderResults();};div.onblur=done;}if(action==='undo'){term[field]=term[field==='native'?'originalNative':'originalTranslation'];prepareSearchTerm(term);rebuildSearchData();await saveTermBook();renderResults();}if(action==='save'){await saveTermBook();showToast('Saved in local library');}$('cellMenu').hidden=true;};
searchInput.addEventListener('input',()=>showSuggestions(searchInput.value));
function showSuggestions(){ $('suggestions').hidden=true; }
async function exportWorkbook(bookId){const book=workbooks.find(item=>item.id===bookId)||workbooks[0];if(!book)return;const zip=await JSZip.loadAsync(book.buffer);const parser=new DOMParser(),serializer=new XMLSerializer(),groups={};book.terms.forEach(t=>(groups[t.sheetPath]??=[]).push(t));for(const [path,items] of Object.entries(groups)){const doc=parser.parseFromString(await zip.file(path).async('text'),'application/xml'),data=doc.querySelector('sheetData');for(const term of items){let row=term.row&&doc.querySelector(`row[r="${term.row}"]`);if(!row){const nums=[...data.querySelectorAll('row')].map(r=>Number(r.getAttribute('r')));term.row=Math.max(...nums,0)+1;row=doc.createElementNS(doc.documentElement.namespaceURI,'row');row.setAttribute('r',term.row);data.append(row);}setInlineCell(doc,row,term.nativeIndex,term.row,term.native);setInlineCell(doc,row,term.translationIndex,term.row,term.translation);}zip.file(path,serializer.serializeToString(doc));}const blob=await zip.generateAsync({type:'blob'}),a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download=book.name.replace(/\.xlsx?$/i,'')+'_updated.xlsx';a.click();setTimeout(()=>URL.revokeObjectURL(a.href),1000);showToast('Updated Excel exported');}
function setInlineCell(doc,row,index,rowNo,value){const ref=`${columnLetters(index)}${rowNo}`;let cell=[...row.querySelectorAll('c')].find(c=>c.getAttribute('r')===ref);if(!cell){cell=doc.createElementNS(doc.documentElement.namespaceURI,'c');cell.setAttribute('r',ref);row.append(cell);}cell.setAttribute('t','inlineStr');cell.replaceChildren();const is=doc.createElementNS(doc.documentElement.namespaceURI,'is'),t=doc.createElementNS(doc.documentElement.namespaceURI,'t');t.textContent=value;if(/^\s|\s$/.test(value))t.setAttribute('xml:space','preserve');is.append(t);cell.append(is);}
function columnLetters(index){let out='';for(let n=index+1;n;n=Math.floor((n-1)/26))out=String.fromCharCode(65+(n-1)%26)+out;return out;}
restoreLibrary();
