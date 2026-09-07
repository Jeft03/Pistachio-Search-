/* Runs only in the browser. No fetch, server, or workbook write is used. */
const $ = (id) => document.getElementById(id);
const dropZone = $('dropZone'), fileInput = $('fileInput'), searchInput = $('searchInput');
let terms = [], clickTimer, searchTimer = null, currentFile = null, activeCell = null;
const openCCReady = Boolean(window.OpenCC && OpenCC.Converter);
const convertTraditionalToSimplified = openCCReady ? OpenCC.Converter({from:'tw', to:'cn'}) : value => value;
const convertSimplifiedToTraditional = openCCReady ? OpenCC.Converter({from:'cn', to:'tw'}) : value => value;

dropZone.addEventListener('click', () => fileInput.click());
dropZone.addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); fileInput.click(); } });
fileInput.addEventListener('change', e => e.target.files[0] && importWorkbook(e.target.files[0]));
['dragenter','dragover'].forEach(type => dropZone.addEventListener(type, e => { e.preventDefault(); dropZone.classList.add('drag'); }));
['dragleave','drop'].forEach(type => dropZone.addEventListener(type, e => { e.preventDefault(); dropZone.classList.remove('drag'); }));
dropZone.addEventListener('drop', e => { const file = [...e.dataTransfer.files].find(f => /\.(xlsx|xlsm|xlsb)$/i.test(f.name)); if (file) importWorkbook(file); else setStatus('Please choose an Excel .xlsx, .xlsm, or .xlsb file.', 'error'); });
searchInput.addEventListener('input', scheduleRenderResults);
setUpSheetColumnResize();

async function importWorkbook(file) { return loadWorkbook(file.name, await file.arrayBuffer()); }
async function loadWorkbook(name, buffer) {
  if (!window.JSZip) return setStatus('The local Excel reader is missing. Keep the vendor folder beside this page.', 'error');
  try {
    setStatus(`Reading ${name} locally…`);
    const zip = await JSZip.loadAsync(buffer);
    const parsed = await readXlsx(zip);
    terms = prepareSearchTerms(parsed.terms);
    if (!terms.length) throw new Error('No sheet with both a Native column and a Translation/Approved Translation column was found.');
    $('searchArea').hidden = $('resultsSection').hidden = false;
    $('fallback').hidden = true;
    currentFile = {name, buffer}; await saveRecent();
    setStatus(`${name}: ${terms.length.toLocaleString()} terms read from ${parsed.usedSheets} sheet${parsed.usedSheets === 1 ? '' : 's'}. Every workbook sheet was read.${openCCReady ? ' Traditional/Simplified matching is ready.' : ' Warning: Traditional/Simplified converter did not load; keep the vendor folder with this page.'}`, openCCReady ? 'good' : 'error');
    searchInput.value = ''; searchInput.focus(); renderResults();
  } catch (error) { console.error(error); setStatus(`Could not read this workbook: ${error.message}`, 'error'); }
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
  const candidates = query ? terms.filter(term => isCandidate(term, tokenData)) : [];
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
    tr.append(cell(term.sheet), copyCell(term, 'native', term.nativeMatch.ranges, exactNative), copyCell(term, 'translation', term.translationMatch.ranges));
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
function copyCell(term, field, ranges, exactNative = false) {
  const text = term[field];
  const td = document.createElement('td'), div = document.createElement('div'); div.className = 'copy-cell'; div.dataset.original = text;
  div.innerHTML = highlightRanges(text, ranges);
  if (exactNative) { const tick = document.createElement('img'); tick.className = 'exact-tick'; tick.src = 'assets/exact-match-tick.png'; tick.alt = 'Exact match'; tick.title = 'Exact match'; div.append(' ', tick); }
  div.addEventListener('click', () => { clearTimeout(clickTimer); clickTimer = setTimeout(() => copyText(text), 190); });
  div.addEventListener('dblclick', e => { e.preventDefault(); clearTimeout(clickTimer); temporaryEdit(div, text); });
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
function temporaryEdit(div, original) { div.classList.add('selecting'); div.contentEditable = 'true'; div.textContent = original; div.focus(); const range = document.createRange(); range.selectNodeContents(div); getSelection().removeAllRanges(); getSelection().addRange(range); const reset = () => { div.contentEditable = 'false'; div.classList.remove('selecting'); renderResults(); }; div.addEventListener('blur', reset, {once:true}); div.addEventListener('keydown', e => { if (e.key === 'Escape') { e.preventDefault(); div.blur(); } }, {once:true}); }
async function copyText(text) { try { await navigator.clipboard.writeText(text); showToast('Copied'); } catch { showToast('Select the text and copy manually'); } }
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

/* Local library, editing and export controls. Nothing is sent to a server. */
const DB_NAME = 'pistachio-search-library';
function db() { return new Promise((resolve, reject) => { const r = indexedDB.open(DB_NAME, 1); r.onupgradeneeded = () => r.result.createObjectStore('books', {keyPath:'name'}); r.onsuccess = () => resolve(r.result); r.onerror = () => reject(r.error); }); }
async function saveRecent() { if (!currentFile) return; const d = await db(); const tx = d.transaction('books','readwrite'); tx.objectStore('books').put({name:currentFile.name,buffer:currentFile.buffer,terms,saved:Date.now()}); await new Promise((resolve,reject)=>{tx.oncomplete=resolve;tx.onerror=()=>reject(tx.error)}); }
async function recentBooks() { const d = await db(); const tx=d.transaction('books','readonly'); const r=tx.objectStore('books').getAll(); return await new Promise((resolve,reject)=>{r.onsuccess=()=>resolve(r.result.sort((a,b)=>b.saved-a.saved));r.onerror=()=>reject(r.error)}); }
async function renderRecentFiles() { const holder=$('recentFiles'); holder.textContent=''; for (const book of await recentBooks()) { const b=document.createElement('button'); b.type='button'; b.textContent=book.name; b.onclick=async()=>{await loadWorkbook(book.name,book.buffer);if(book.terms){terms=prepareSearchTerms(book.terms);renderResults();}}; holder.append(b); } }
function showCellMenu(x,y){const m=$('cellMenu');m.hidden=false;m.style.left=`${x}px`;m.style.top=`${y}px`;}
document.addEventListener('click',event=>{if(!event.target.closest('#cellMenu'))$('cellMenu').hidden=true;});
$('cellMenu').onclick=async event=>{const action=event.target.dataset.action;if(!action||!activeCell)return;const {term,field,div}=activeCell;if(action==='edit'){div.contentEditable='true';div.classList.add('selecting');div.focus();const done=async()=>{term[field]=div.textContent.trim()||term[field];prepareSearchTerm(term);div.contentEditable='false';div.classList.remove('selecting');await saveRecent();renderResults();};div.onblur=done;}if(action==='undo'){term[field]=term[field==='native'?'originalNative':'originalTranslation'];prepareSearchTerm(term);await saveRecent();renderResults();}if(action==='save'){await saveRecent();showToast('Saved in local library');}$('cellMenu').hidden=true;};
searchInput.addEventListener('input',()=>showSuggestions(searchInput.value));
function showSuggestions(){ $('suggestions').hidden=true; }
async function exportWorkbook(){if(!currentFile)return;const zip=await JSZip.loadAsync(currentFile.buffer);const parser=new DOMParser(),serializer=new XMLSerializer(),groups={};terms.forEach(t=>(groups[t.sheetPath]??=[]).push(t));for(const [path,items] of Object.entries(groups)){const doc=parser.parseFromString(await zip.file(path).async('text'),'application/xml'),data=doc.querySelector('sheetData');for(const term of items){let row=term.row&&doc.querySelector(`row[r="${term.row}"]`);if(!row){const nums=[...data.querySelectorAll('row')].map(r=>Number(r.getAttribute('r')));term.row=Math.max(...nums,0)+1;row=doc.createElementNS(doc.documentElement.namespaceURI,'row');row.setAttribute('r',term.row);data.append(row);}setInlineCell(doc,row,term.nativeIndex,term.row,term.native);setInlineCell(doc,row,term.translationIndex,term.row,term.translation);}zip.file(path,serializer.serializeToString(doc));}const blob=await zip.generateAsync({type:'blob'}),a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download=currentFile.name.replace(/\.xlsx?$/i,'')+'_updated.xlsx';a.click();setTimeout(()=>URL.revokeObjectURL(a.href),1000);showToast('Updated Excel exported');}
function setInlineCell(doc,row,index,rowNo,value){const ref=`${columnLetters(index)}${rowNo}`;let cell=[...row.querySelectorAll('c')].find(c=>c.getAttribute('r')===ref);if(!cell){cell=doc.createElementNS(doc.documentElement.namespaceURI,'c');cell.setAttribute('r',ref);row.append(cell);}cell.setAttribute('t','inlineStr');cell.replaceChildren();const is=doc.createElementNS(doc.documentElement.namespaceURI,'is'),t=doc.createElementNS(doc.documentElement.namespaceURI,'t');t.textContent=value;if(/^\s|\s$/.test(value))t.setAttribute('xml:space','preserve');is.append(t);cell.append(is);}
function columnLetters(index){let out='';for(let n=index+1;n;n=Math.floor((n-1)/26))out=String.fromCharCode(65+(n-1)%26)+out;return out;}
