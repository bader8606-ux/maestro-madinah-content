/** Source-faithful XLSX model, safe calculation, editing and OOXML export.
 * Workbook data is supplied at runtime and is never bundled with this module.
 * Supported formulas: references, ranges, SUM, arithmetic and parentheses.
 */
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';

const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const MAX_ARCHIVE_BYTES = 10 * 1024 * 1024;
const MAX_UNCOMPRESSED_BYTES = 50 * 1024 * 1024;
const MAX_DATA_ROW = 20000;
const MAX_FORMULA_LENGTH = 1000;
const REF_PATTERN = "(?:(?:'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_.]*)!)?\\$?[A-Za-z]{1,3}\\$?[1-9][0-9]*";
const REF_PARTS = /^(?:(?<sheet>'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_.]*)!)?(?<ca>\$?)(?<col>[A-Za-z]{1,3})(?<ra>\$?)(?<row>[1-9][0-9]*)$/;
const TOKEN_RE = new RegExp(`\\s*(?:(${REF_PATTERN})|((?:\\d+(?:\\.\\d*)?|\\.\\d+)(?:[eE][+-]?\\d+)?)|([A-Za-z_][A-Za-z0-9_]*)|([+*/(),:\\-]))`, 'y');
const TOTAL_RE = /^=\s*SUM\s*\(\s*\$?([AC])\$?(\d+)\s*:\s*\$?\1\$?(\d+)\s*\)\s*$/i;
const METRIC_COLS = { ratio: 'D', seconds: 'E', count: 'F', dailyHours: 'G', monthlyHours: 'I' };
const clone = value => structuredClone(value);
const isNumber = value => typeof value === 'number';
const attr = (element, name, fallback = null) => element.hasAttribute(name) ? element.getAttribute(name) : fallback;
const children = (element, name, ns = MAIN) => Array.from(element?.childNodes || []).filter(node => node.nodeType === 1 && node.localName === name && node.namespaceURI === ns);
const child = (element, name, ns = MAIN) => children(element, name, ns)[0] || null;
const descendants = (element, name, ns = MAIN) => Array.from(element.getElementsByTagNameNS(ns, name));
const attrs = element => Object.fromEntries(Array.from(element.attributes || []).map(item => [item.name, item.value]));
const serializer = new XMLSerializer();

class CalcError extends Error {
  constructor(code, detail = '') { super(code); this.code = code; this.detail = detail; }
}

export function colNumber(value) {
  let result = 0;
  for (const character of value.toUpperCase()) result = result * 26 + character.charCodeAt(0) - 64;
  return result;
}

export function colLabel(value) {
  let result = '';
  while (value) { const remainder = (value - 1) % 26; value = Math.floor((value - 1) / 26); result = String.fromCharCode(65 + remainder) + result; }
  return result;
}

function splitRef(reference) {
  const match = typeof reference === 'string' && REF_PARTS.exec(reference);
  if (!match) throw new Error('مرجع خلية غير مدعوم: ' + String(reference));
  return match.groups;
}
function plainRef(reference) { const m = splitRef(reference); return m.col.toUpperCase() + m.row; }
function sheetName(reference, fallback) {
  const name = splitRef(reference).sheet;
  return name ? (name.startsWith("'") ? name.slice(1, -1).replaceAll("''", "'") : name) : fallback;
}
function refSort(a, b) { const x = splitRef(a), y = splitRef(b); return Number(x.row) - Number(y.row) || colNumber(x.col) - colNumber(y.col); }
function rowOf(reference) { return Number(splitRef(reference).row); }

function tokens(formula) {
  const text = formula.startsWith('=') ? formula.slice(1) : formula;
  if (text.length > MAX_FORMULA_LENGTH) throw new Error('معادلة أطول من الحد التقني المدعوم.');
  const result = [];
  let position = 0;
  while (position < text.length) {
    if (!text.slice(position).trim()) break;
    TOKEN_RE.lastIndex = position;
    const match = TOKEN_RE.exec(text);
    if (!match) throw new Error('صيغة غير مدعومة في المعادلة: ' + formula);
    const i = [1, 2, 3, 4].find(index => match[index] !== undefined);
    const kind = ['ref', 'number', 'ident', 'op'][i - 1];
    const value = match[i];
    const start = TOKEN_RE.lastIndex - value.length;
    result.push([kind, value, start, TOKEN_RE.lastIndex]);
    position = TOKEN_RE.lastIndex;
    if (result.length > 500) throw new Error('عدد عناصر المعادلة يتجاوز الحد التقني المدعوم.');
  }
  return result;
}

function replaceRefs(formula, replace) {
  const parsed = tokens(formula);
  if (!parsed.length) return formula;
  const text = formula.startsWith('=') ? formula.slice(1) : formula;
  let result = '', position = 0;
  for (const [kind, value, start, stop] of parsed) {
    result += text.slice(position, start) + (kind === 'ref' ? replace(value) : text.slice(start, stop));
    position = stop;
  }
  return (formula.startsWith('=') ? '=' : '') + result + text.slice(position);
}

export function copyFormula(formula, rowDelta = 0, colDelta = 0) {
  return replaceRefs(formula, reference => {
    const m = splitRef(reference);
    const row = Number(m.row) + (m.ra ? 0 : rowDelta);
    const col = colNumber(m.col) + (m.ca ? 0 : colDelta);
    if (!(row >= 1 && row <= 1048576 && col >= 1 && col <= 16384)) throw new Error('لا يمكن تمديد المعادلة خارج نطاق Excel.');
    return (m.sheet ? m.sheet + '!' : '') + m.ca + colLabel(col) + m.ra + row;
  });
}
function shiftFormula(formula, formulaSheet, insertedSheet, at) {
  return replaceRefs(formula, reference => {
    if (sheetName(reference, formulaSheet) !== insertedSheet) return reference;
    const m = splitRef(reference);
    if (Number(m.row) < at) return reference;
    return (m.sheet ? m.sheet + '!' : '') + m.ca + m.col + m.ra + (Number(m.row) + 1);
  });
}

class FormulaParser {
  constructor(formula) { this.formula = formula; this.tokens = tokens(formula); this.position = 0; this.depth = 0; }
  peek(value) {
    if (this.position >= this.tokens.length) return value === undefined ? null : false;
    const token = this.tokens[this.position];
    return value === undefined ? token : token[1].toUpperCase() === value;
  }
  take(expected) {
    const token = this.peek();
    if (token === null || (expected !== undefined && token[1].toUpperCase() !== expected)) throw new Error('بنية معادلة غير مدعومة: ' + this.formula);
    this.position += 1;
    return token;
  }
  parse() {
    if (!this.tokens.length) throw new Error('معادلة فارغة غير صالحة.');
    const result = this.expression();
    if (this.peek() !== null) throw new Error('بنية معادلة غير مدعومة: ' + this.formula);
    return result;
  }
  expression() {
    let node = this.term();
    while (this.peek('+') || this.peek('-')) node = ['binary', this.take()[1], node, this.term()];
    return node;
  }
  term() {
    let node = this.factor();
    while (this.peek('*') || this.peek('/')) node = ['binary', this.take()[1], node, this.factor()];
    return node;
  }
  factor() {
    this.depth += 1;
    if (this.depth > 100) throw new Error('تداخل المعادلة يتجاوز الحد التقني المدعوم.');
    try {
      const token = this.peek();
      if (token === null) throw new Error('معادلة غير مكتملة: ' + this.formula);
      const [kind, value] = token;
      if (value === '+' || value === '-') { this.take(); return ['unary', value, this.factor()]; }
      if (value === '(') { this.take('('); const node = this.expression(); this.take(')'); return node; }
      if (kind === 'number') {
        this.take(); const number = Number(value);
        if (!Number.isFinite(number)) throw new Error('رقم غير محدود في المعادلة.');
        return ['number', number];
      }
      if (kind === 'ref') {
        this.take();
        if (this.peek(':')) { this.take(':'); const second = this.take(); if (second[0] !== 'ref') throw new Error('نطاق خلايا غير مدعوم.'); return ['range', value, second[1]]; }
        return ['ref', value];
      }
      if (kind === 'ident' && value.toUpperCase() === 'SUM') {
        this.take(); this.take('('); const args = [];
        if (!this.peek(')')) { args.push(this.expression()); while (this.peek(',')) { this.take(','); args.push(this.expression()); } }
        this.take(')'); return ['sum', args];
      }
      throw new Error('الدالة أو الصيغة غير مدعومة: ' + this.formula);
    } finally { this.depth -= 1; }
  }
}

function asNumber(value) {
  if (value === null || value === undefined) return 0;
  if (typeof value === 'boolean') return Number(value);
  if (isNumber(value)) { if (Number.isFinite(value)) return value; throw new CalcError('#NUM!'); }
  if (typeof value === 'string' && value.trim() && Number.isFinite(Number(value.trim()))) return Number(value.trim());
  throw new CalcError('#VALUE!');
}

class Calculator {
  constructor(model) { this.sheets = new Map(model.sheets.map(sheet => [sheet.name, sheet])); this.memo = new Map(); this.errors = new Map(); this.active = new Set(); this.trees = new Map(); }
  cell(name, reference) {
    const ref = plainRef(reference), key = name + '\0' + ref;
    if (this.memo.has(key)) { if (this.errors.has(key)) throw new CalcError(this.errors.get(key)); return this.memo.get(key); }
    if (!this.sheets.has(name)) throw new CalcError('#REF!');
    if (this.active.has(key) || this.active.size > 100) throw new CalcError('#REF!', 'مرجع دائري أو تداخل طويل');
    const cell = this.sheets.get(name).cells[ref];
    if (!cell) { this.memo.set(key, null); return null; }
    if (!cell.formula) {
      if (cell.attrs?.t === 'e' && cell.value) { this.errors.set(key, String(cell.value)); this.memo.set(key, null); throw new CalcError(String(cell.value)); }
      const value = cell.value ?? null; this.memo.set(key, value); return value;
    }
    this.active.add(key);
    try {
      if (!this.trees.has(cell.formula)) this.trees.set(cell.formula, new FormulaParser(cell.formula).parse());
      const result = this.node(this.trees.get(cell.formula), name);
      if (Array.isArray(result)) throw new CalcError('#VALUE!');
      if (isNumber(result) && !Number.isFinite(result)) throw new CalcError('#NUM!');
      this.memo.set(key, result); return result;
    } catch (error) {
      if (error instanceof CalcError) { this.memo.set(key, null); this.errors.set(key, error.code); }
      throw error;
    } finally { this.active.delete(key); }
  }
  node(node, currentSheet) {
    const kind = node[0];
    if (kind === 'number') return node[1];
    if (kind === 'ref') return this.cell(sheetName(node[1], currentSheet), node[1]);
    if (kind === 'range') {
      const left = splitRef(node[1]), right = splitRef(node[2]);
      const name = sheetName(node[1], currentSheet), rightName = sheetName(node[2], name);
      if (name !== rightName) throw new CalcError('#REF!');
      const r1 = Math.min(Number(left.row), Number(right.row)), r2 = Math.max(Number(left.row), Number(right.row));
      const c1 = Math.min(colNumber(left.col), colNumber(right.col)), c2 = Math.max(colNumber(left.col), colNumber(right.col));
      if ((r2 - r1 + 1) * (c2 - c1 + 1) > 50000) throw new Error('نطاق المعادلة أكبر من الحد التقني المدعوم.');
      const result = [];
      for (let row = r1; row <= r2; row++) for (let col = c1; col <= c2; col++) result.push(this.cell(name, colLabel(col) + row));
      return result;
    }
    if (kind === 'unary') { const value = asNumber(this.node(node[2], currentSheet)); return node[1] === '-' ? -value : value; }
    if (kind === 'binary') {
      const left = asNumber(this.node(node[2], currentSheet)), right = asNumber(this.node(node[3], currentSheet));
      if (node[1] === '+') return left + right;
      if (node[1] === '-') return left - right;
      if (node[1] === '*') return left * right;
      if (right === 0) throw new CalcError('#DIV/0!');
      return left / right;
    }
    if (kind === 'sum') {
      let total = 0;
      for (const argument of node[1]) {
        const value = this.node(argument, currentSheet);
        for (const item of Array.isArray(value) ? value : [value]) if (isNumber(item)) total += item;
      }
      return total;
    }
    throw new Error('نوع حساب غير مدعوم.');
  }
  safe(name, reference) { try { return [this.cell(name, reference), null]; } catch (error) { if (error instanceof CalcError) return [null, error.code]; throw error; } }
}

function xml(data) {
  const text = typeof data === 'string' ? data : strFromU8(data);
  if (/<!DOCTYPE|<!ENTITY/i.test(text)) throw new Error('تعريفات XML الخارجية غير مسموحة.');
  try {
    const errors = [];
    // xmldom's errorHandler is required: its default parser can recover malformed
    // XML. Recovery is unsuitable when preserving an imported workbook exactly.
    const report = message => errors.push(message);
    const doc = new DOMParser({ errorHandler: { warning: report, error: report, fatalError: report } }).parseFromString(text, 'application/xml');
    if (!doc.documentElement || errors.length || descendants(doc, 'parsererror', '*').length) throw new Error('XML');
    return doc.documentElement;
  } catch { throw new Error('ملف XML داخل المصنف غير صالح.'); }
}

function archive(data) {
  if (data instanceof ArrayBuffer) data = new Uint8Array(data);
  if (!(data instanceof Uint8Array) || data.length > MAX_ARCHIVE_BYTES) throw new Error('حجم الملف يتجاوز 10 ميغابايت أو تنسيقه غير صالح.');
  try {
    // Inspect the central directory before decompression to bound memory usage.
    const reader = new DataView(data.buffer, data.byteOffset, data.byteLength);
    let end = -1;
    for (let p = data.length - 22; p >= Math.max(0, data.length - 65557); p--) if (reader.getUint32(p, true) === 0x06054b50 && p + 22 + reader.getUint16(p + 20, true) === data.length) { end = p; break; }
    if (end < 0 || reader.getUint16(end + 4, true) || reader.getUint16(end + 6, true)) throw new Error('الملف ليس مصنف Excel بصيغة XLSX صالحًا.');
    const count = reader.getUint16(end + 10, true), directory = reader.getUint32(end + 16, true);
    if (count > 300 || count === 65535 || directory === 0xffffffff) throw new Error('بنية حزمة Excel غير صالحة أو أكبر من المدعوم.');
    const names = new Set(); let total = 0, position = directory;
    for (let i = 0; i < count; i++) {
      if (position + 46 > data.length || reader.getUint32(position, true) !== 0x02014b50) throw new Error('بنية حزمة Excel غير صالحة أو أكبر من المدعوم.');
      const flags = reader.getUint16(position + 8, true), size = reader.getUint32(position + 24, true);
      const nameLength = reader.getUint16(position + 28, true), extraLength = reader.getUint16(position + 30, true), commentLength = reader.getUint16(position + 32, true);
      const name = strFromU8(data.subarray(position + 46, position + 46 + nameLength));
      if (names.has(name)) throw new Error('بنية حزمة Excel غير صالحة أو أكبر من المدعوم.');
      names.add(name); total += size;
      if (total > MAX_UNCOMPRESSED_BYTES) throw new Error('المحتوى غير المضغوط أكبر من الحد التقني المدعوم.');
      if (flags & 1) throw new Error('المصنفات المشفرة غير مدعومة.');
      if (name.includes('vbaProject') || name.startsWith('xl/externalLinks/')) throw new Error('الماكرو والارتباطات بمصنفات خارجية غير مدعومة.');
      position += 46 + nameLength + extraLength + commentLength;
    }
    for (const required of ['xl/workbook.xml', 'xl/_rels/workbook.xml.rels', '[Content_Types].xml']) if (!names.has(required)) throw new Error('أحد أجزاء مصنف Excel الأساسية مفقود.');
    const files = unzipSync(data);
    if (Object.values(files).reduce((sum, value) => sum + value.length, 0) > MAX_UNCOMPRESSED_BYTES) throw new Error('المحتوى غير المضغوط أكبر من الحد التقني المدعوم.');
    return files;
  } catch (error) {
    if (error instanceof RangeError || !error.message || !/[\u0600-\u06ff]/.test(error.message)) throw new Error('الملف ليس مصنف Excel بصيغة XLSX صالحًا.');
    throw error;
  }
}

function typedValue(element, strings) {
  const kind = attr(element, 't'), valueNode = child(element, 'v'), value = valueNode ? valueNode.textContent : null;
  if (kind === 's') {
    if (value === null) return null;
    if (!/^\d+$/.test(value) || Number(value) >= strings.length) throw new Error('فهرس نص داخل المصنف غير صالح.');
    return strings[Number(value)];
  }
  if (kind === 'inlineStr') return descendants(element, 't').map(node => node.textContent).join('');
  if (kind === 'b') return value === null ? null : value === '1';
  if (['str', 'e', 'd'].includes(kind)) return value;
  if (value === null) return null;
  if (!value.trim() || !Number.isFinite(Number(value))) throw new Error('قيمة رقمية غير صالحة داخل المصنف.');
  return Number(value);
}

function makeGroups(sheet) {
  const sumColumn = sheet.name === 'Sheet1' ? 'A' : 'C';
  const totalRows = [];
  for (const [reference, cell] of Object.entries(sheet.cells)) {
    const match = TOTAL_RE.exec(cell.formula || '');
    if (!match || match[1].toUpperCase() !== sumColumn) continue;
    const total = rowOf(reference), start = Number(match[2]), end = Number(match[3]);
    if (reference === sumColumn + total && start <= end && end < total) totalRows.push([total, start, end]);
  }
  totalRows.sort((a, b) => a[0] - b[0]);
  const groups = []; let priorTotal = 0;
  for (const [total, start, end] of totalRows) {
    if (start <= priorTotal) throw new Error('نطاقات مجموعات متداخلة لا يمكن استيرادها بأمان.');
    let titleRef = null;
    const titleColumn = sheet.name === 'Sheet1' ? 'A' : 'B';
    for (let row = priorTotal + 1; row < start; row++) {
      const value = sheet.cells[titleColumn + row]?.value;
      if (typeof value === 'string' && value.trim() && value.trim() !== 'اسم المحتوى') titleRef = titleColumn + row;
    }
    const notes = [], records = [];
    for (let row = start; row <= end; row++) {
      const duration = sheet.cells[sumColumn + row]?.value ?? null, name = sheet.cells['B' + row]?.value ?? null;
      const extra = sumColumn === 'A' ? sheet.cells['C' + row]?.value ?? null : null;
      if (sumColumn === 'A' && typeof duration === 'string' && !name) { notes.push('A' + row); continue; }
      if (name !== null || isNumber(duration) || extra !== null) records.push({ id: sheet.name + '-r' + row, row, deleted: false });
    }
    groups.push({ id: sheet.name + '-g' + (groups.length + 1), startRow: start, endRow: end, totalRow: total, titleRef,
      headerRow: sumColumn === 'C' ? start - 1 : null, noteRefs: notes, records });
    priorTotal = total;
  }
  if (!groups.length) throw new Error('لم توجد مجموعات ذات إجماليات SUM في ' + sheet.name + '.');
  return groups;
}

async function sha256(data) {
  const result = await globalThis.crypto.subtle.digest('SHA-256', data);
  return Array.from(new Uint8Array(result), value => value.toString(16).padStart(2, '0')).join('');
}

function normalizePath(path) {
  const parts = [];
  for (const part of path.split('/')) { if (part === '..') parts.pop(); else if (part && part !== '.') parts.push(part); }
  return parts.join('/');
}

export async function parseXlsx(data, filename) {
  if (typeof filename !== 'string' || !filename.toLowerCase().endsWith('.xlsx')) throw new Error('يرجى اختيار ملف Excel بصيغة .xlsx.');
  if (data instanceof ArrayBuffer) data = new Uint8Array(data);
  const files = archive(data);
  const strings = files['xl/sharedStrings.xml'] ? children(xml(files['xl/sharedStrings.xml']), 'si').map(item => descendants(item, 't').map(node => node.textContent).join('')) : [];
  const workbook = xml(files['xl/workbook.xml']);
  const relationships = Object.fromEntries(Array.from(xml(files['xl/_rels/workbook.xml.rels']).childNodes).filter(node => node.nodeType === 1).map(entry => [attr(entry, 'Id'), attr(entry, 'Target')]));
  const model = { version: 1, sourceName: filename.replaceAll('\\', '/').split('/').at(-1), sourceHash: await sha256(data), sheets: [] };
  for (const descriptor of children(child(workbook, 'sheets'), 'sheet')) {
    const name = attr(descriptor, 'name');
    const target = relationships[descriptor.getAttributeNS(REL, 'id')] || '';
    const part = normalizePath(target.startsWith('/') ? target.slice(1) : 'xl/' + target);
    if (!part.startsWith('xl/worksheets/') || !Object.hasOwn(files, part)) throw new Error('رابط ورقة Excel غير صالح.');
    const root = xml(files[part]), cells = {}, rowAttrs = {}, shared = {}, followers = [];
    let maxRow = 0, maxColumn = 0;
    const sheetData = child(root, 'sheetData');
    if (!sheetData) throw new Error('بيانات ورقة Excel مفقودة.');
    for (const row of children(sheetData, 'row')) {
      const rowNumber = Number(attr(row, 'r', '0'));
      if (!Number.isInteger(rowNumber) || rowNumber < 1 || rowNumber > MAX_DATA_ROW) throw new Error('عدد صفوف المصنف يتجاوز الحد التقني المدعوم (20000).');
      maxRow = Math.max(maxRow, rowNumber); rowAttrs[String(rowNumber)] = attrs(row);
      for (const element of children(row, 'c')) {
        const reference = attr(element, 'r', ''), match = splitRef(reference);
        if (Number(match.row) !== rowNumber || colNumber(match.col) > 256) throw new Error('نطاق خلايا غير مدعوم.');
        maxColumn = Math.max(maxColumn, colNumber(match.col));
        const formulaElement = child(element, 'f'), value = typedValue(element, strings);
        let formula = null;
        if (formulaElement) {
          if (![null, 'normal', 'shared'].includes(attr(formulaElement, 't'))) throw new Error('المعادلات المصفوفية وجداول الحساسية غير مدعومة في هذا التطبيق.');
          formula = formulaElement.textContent ? '=' + formulaElement.textContent.replace(/^=+/, '') : null;
          if (attr(formulaElement, 't') === 'shared') {
            const index = attr(formulaElement, 'si');
            if (formula) shared[index] = [reference, formula]; else followers.push([reference, index]);
          }
        }
        cells[reference] = { value, cached: value, formula, attrs: attrs(element), xml: serializer.serializeToString(element), originalRef: reference, dirty: false };
      }
    }
    for (const [reference, index] of followers) {
      if (!Object.hasOwn(shared, index)) throw new Error('معادلة مشتركة بلا خلية مرجعية.');
      const [anchor, formula] = shared[index], a = splitRef(anchor), b = splitRef(reference);
      cells[reference].formula = copyFormula(formula, Number(b.row) - Number(a.row), colNumber(b.col) - colNumber(a.col));
    }
    for (const cell of Object.values(cells)) if (cell.formula) new FormulaParser(cell.formula).parse();
    model.sheets.push({ name, part, cells, rowAttrs, maxRow, maxColumn, insertions: [], merges: child(root, 'mergeCells') ? children(child(root, 'mergeCells'), 'mergeCell').map(item => attr(item, 'ref')) : [] });
  }
  if (JSON.stringify(model.sheets.map(sheet => sheet.name)) !== JSON.stringify(['Sheet1', 'Sheet2'])) throw new Error('يرجى استيراد المصنف بالهيكل نفسه وورقتي Sheet1 وSheet2 بهذا الترتيب.');
  for (const sheet of model.sheets) sheet.groups = makeGroups(sheet);
  const calculator = new Calculator(model);
  for (const sheet of model.sheets) for (const [reference, cell] of Object.entries(sheet.cells)) if (cell.formula) calculator.safe(sheet.name, reference);
  return model;
}

function inputRefs(sheet, row) { return { name: 'B' + row, duration: (sheet.name === 'Sheet1' ? 'A' : 'C') + row, extra: sheet.name === 'Sheet1' ? 'C' + row : null, days: sheet.name === 'Sheet2' ? 'H' + row : null }; }
const sixDigits = number => Number(number.toPrecision(6)).toString();

export function view(model) {
  const calculator = new Calculator(model), result = { sheets: [], warnings: [], cellCount: 0, formulaCount: 0 };
  for (const sheet of model.sheets) {
    const name = sheet.name, output = { name, groups: [], rows: [] }, editableRefs = new Set();
    for (const group of sheet.groups) {
      const start = group.startRow, total = group.totalRow, rangeColumn = name === 'Sheet1' ? 'A' : 'C';
      const sourceRange = rangeColumn + start + ':' + rangeColumn + total;
      const title = group.titleRef ? sheet.cells[group.titleRef]?.value : null;
      const hasTitle = typeof title === 'string' && Boolean(title.trim());
      const publicGroup = { id: group.id, sheet: name, title: hasTitle ? title : 'مجموعة دون عنوان · ' + sourceRange, hasTitle, sourceRange, records: [], totals: {}, totalRefs: {}, notes: [] };
      const warn = (code, message, reference) => result.warnings.push({ code, message, sheet: name, ref: reference, groupId: group.id });
      if (!hasTitle) warn('missing_group_title', 'لا يوجد عنوان لهذه المجموعة في المصنف الأصلي.', sourceRange);
      for (const record of group.records) {
        if (record.deleted) continue;
        const row = record.row, refs = inputRefs(sheet, row);
        const values = Object.fromEntries(Object.entries(refs).map(([key, reference]) => [key, reference ? calculator.safe(name, reference)[0] : null]));
        for (const reference of Object.values(refs)) if (reference && !sheet.cells[reference]?.formula) editableRefs.add(reference);
        const publicRecord = { id: record.id, sourceRow: row, ...values, inputRefs: refs, metrics: {}, formulas: {}, errors: {}, missingOutputs: [] };
        if (typeof values.name !== 'string' || !values.name.trim()) warn('missing_name', 'اسم المحتوى غير مذكور في ' + refs.name + '.', refs.name);
        if (values.duration === null) warn('missing_duration', 'مدة المحتوى غير مذكورة في ' + refs.duration + '.', refs.duration);
        else if (!isNumber(values.duration) || values.duration <= 0) warn('invalid_duration', 'مدة المحتوى تحتاج قيمة رقمية أكبر من صفر.', refs.duration);
        for (const [key, column] of Object.entries(METRIC_COLS)) {
          const reference = column + row, cell = name === 'Sheet2' ? sheet.cells[reference] || {} : {};
          const [number, error] = name === 'Sheet2' ? calculator.safe(name, reference) : [null, null];
          publicRecord.metrics[key] = number; publicRecord.formulas[key] = cell.formula || null;
          if (error) publicRecord.errors[key] = error;
          if (name === 'Sheet2' && cell.value == null && !cell.formula) publicRecord.missingOutputs.push(key);
        }
        if (name === 'Sheet2' && publicRecord.missingOutputs.length) {
          const absent = publicRecord.missingOutputs.map(key => METRIC_COLS[key] + row).join(', ');
          warn('missing_outputs', 'خلايا حسابية خالية في المصدر: ' + absent + '. لم تُنشأ معادلات بديلة.', 'B' + row);
        }
        if (name === 'Sheet2' && values.days === null) warn('missing_days', 'عدد الأيام غير مذكور في ' + refs.days + '.', refs.days);
        publicGroup.records.push(publicRecord);
      }
      const columns = name === 'Sheet1' ? { duration: 'A' } : { duration: 'C', operatingSeconds: 'D', seconds: 'E', dailyHours: 'G' };
      for (const [key, column] of Object.entries(columns)) { const reference = column + total; publicGroup.totalRefs[key] = reference; publicGroup.totals[key] = calculator.safe(name, reference)[0]; }
      for (const reference of group.noteRefs || []) publicGroup.notes.push({ ref: reference, text: sheet.cells[reference]?.value ?? null });
      if (name === 'Sheet2') {
        if (publicGroup.records.length) {
          const denominator = /\/\s*(\$?[A-Za-z]+\$?\d+)/.exec(publicGroup.records[0].formulas.ratio || '');
          if (denominator) {
            const reference = plainRef(denominator[1]); publicGroup.denominatorRef = reference;
            if (reference !== 'C' + total) warn('shared_denominator', 'نِسَب هذه المجموعة مرتبطة بالمقام ' + reference + ' وليس بإجمالي مدتها C' + total + '. أُبقي منطق المصدر كما هو.', reference);
          }
        }
        const seconds = publicGroup.totals.seconds, operating = publicGroup.totals.operatingSeconds, hours = publicGroup.totals.dailyHours;
        if (isNumber(seconds) && isNumber(operating) && seconds > operating + 1e-8) warn('over_day_capacity', 'إجمالي الثواني يتجاوز مدة التشغيل اليومية المحددة في المصدر.', 'E' + total);
        if (isNumber(seconds) && isNumber(hours) && Math.abs(seconds / 3600 - hours) > 1e-8) warn('incomplete_hour_total', 'إجمالي الثواني يعادل ' + sixDigits(seconds / 3600) + ' ساعة، لكن مجموع عمود الساعات يساوي ' + sixDigits(hours) + ' بسبب الخلايا الخالية في المصدر.', 'G' + total);
      }
      output.groups.push(publicGroup);
    }
    const byRow = new Map();
    for (const reference of Object.keys(sheet.cells).sort(refSort)) {
      const cell = sheet.cells[reference], [value, error] = calculator.safe(name, reference);
      if (cell.formula) result.formulaCount += 1;
      if (cell.formula || cell.value != null) result.cellCount += 1;
      if (error) result.warnings.push({ code: 'formula_error', message: 'خطأ حسابي ' + error, sheet: name, ref: reference, groupId: null });
      const row = rowOf(reference); if (!byRow.has(row)) byRow.set(row, []);
      byRow.get(row).push({ ref: reference, value, formula: cell.formula || null, error, editable: editableRefs.has(reference) && !cell.formula });
    }
    output.rows = Array.from(byRow, ([row, cells]) => ({ row, cells })); result.sheets.push(output);
  }
  return result;
}

export function baselineMismatches(model, tolerance = 1e-8) {
  const calculator = new Calculator(model), mismatches = [];
  for (const sheet of model.sheets) for (const [reference, cell] of Object.entries(sheet.cells)) {
    if (!cell.formula) continue;
    const [actual, error] = calculator.safe(sheet.name, reference), cached = cell.cached;
    const equal = error === null && isNumber(actual) && isNumber(cached) && Math.abs(actual - cached) <= Math.max(tolerance, tolerance * Math.max(Math.abs(actual), Math.abs(cached)));
    if (!equal) mismatches.push({ sheet: sheet.name, ref: reference, cached, calculated: actual, error, formula: cell.formula });
  }
  return mismatches;
}

function shiftRef(reference, at) { return reference.replace(/(\$?)([A-Za-z]{1,3})(\$?)([1-9][0-9]*)/g, (_, ca, col, ra, row) => ca + col + ra + (Number(row) + (Number(row) >= at ? 1 : 0))); }
function rowCells(sheet, row) { return Object.fromEntries(Object.entries(sheet.cells).filter(([ref]) => rowOf(ref) === row).map(([ref, cell]) => [splitRef(ref).col.toUpperCase(), cell])); }

function insertRow(model, sheet, group, templateRow, templateCells = null) {
  const at = group.totalRow, oldEnd = group.endRow;
  if (sheet.maxRow >= MAX_DATA_ROW) throw new Error('وصل المصنف إلى الحد التقني المدعوم للصفوف.');
  const template = clone(templateCells === null ? rowCells(sheet, templateRow) : templateCells);
  const templateAttrs = clone(sheet.rowAttrs[String(templateRow)] || {});
  sheet.cells = Object.fromEntries(Object.entries(sheet.cells).map(([reference, cell]) => [shiftRef(reference, at), cell]));
  for (const otherSheet of model.sheets) {
    for (const cell of Object.values(otherSheet.cells)) if (cell.formula) {
      const updated = shiftFormula(cell.formula, otherSheet.name, sheet.name, at);
      if (updated !== cell.formula) { cell.formula = updated; cell.dirty = true; }
    }
    for (const otherGroup of otherSheet.groups) for (const archived of otherGroup.records) for (const cell of Object.values(archived.deletedCells || {})) if (cell.formula) cell.formula = shiftFormula(cell.formula, otherSheet.name, sheet.name, at);
  }
  sheet.rowAttrs = Object.fromEntries(Object.entries(sheet.rowAttrs).map(([row, attributes]) => [String(Number(row) + (Number(row) >= at ? 1 : 0)), attributes]));
  templateAttrs.r = String(at); sheet.rowAttrs[String(at)] = templateAttrs;
  sheet.merges = (sheet.merges || []).map(reference => shiftRef(reference, at));
  for (const otherGroup of sheet.groups) {
    for (const key of ['startRow', 'endRow', 'totalRow']) if (otherGroup[key] >= at) otherGroup[key] += 1;
    if (otherGroup.headerRow && otherGroup.headerRow >= at) otherGroup.headerRow += 1;
    if (otherGroup.titleRef) otherGroup.titleRef = shiftRef(otherGroup.titleRef, at);
    otherGroup.noteRefs = (otherGroup.noteRefs || []).map(reference => shiftRef(reference, at));
    for (const record of otherGroup.records) if (record.row >= at) record.row += 1;
  }
  group.endRow = at;
  for (const [column, cell] of Object.entries(template)) {
    cell.originalRef = null; cell.dirty = true;
    if (cell.formula) {
      cell.formula = shiftFormula(cell.formula, sheet.name, sheet.name, at);
      cell.formula = copyFormula(cell.formula, at - (templateRow + (templateRow >= at ? 1 : 0)));
    }
    sheet.cells[column + at] = cell;
  }
  for (const [reference, cell] of Object.entries(sheet.cells)) {
    if (rowOf(reference) !== group.totalRow || !cell.formula) continue;
    cell.formula = cell.formula.replace(/(\$?[A-Za-z]+\$?)(\d+):(\$?[A-Za-z]+\$?)(\d+)/g, (match, startCol, startRow, endCol, endRow) => Number(startRow) === group.startRow && Number(endRow) === oldEnd && startCol.replaceAll('$', '') === endCol.replaceAll('$', '') ? startCol + startRow + ':' + endCol + at : match);
    cell.dirty = true;
  }
  sheet.maxRow += 1; sheet.insertions.push(at);
  return at;
}

function validateNumber(value, field, { optional = false, integer = false, positive = false } = {}) {
  if (value === null || value === '') { if (optional) return null; throw new Error('يرجى إدخال ' + field + '.'); }
  if (!isNumber(value) || !Number.isFinite(value)) throw new Error(field + ' يجب أن يكون رقمًا صالحًا.');
  if (positive ? value <= 0 : value < 0) throw new Error(field + (positive ? ' يجب أن يكون أكبر من صفر.' : ' لا يمكن أن يكون سالبًا.'));
  if (integer && !Number.isInteger(value)) throw new Error(field + ' يجب أن يكون عددًا صحيحًا.');
  return value;
}

export function mutate(model, action, groupId, recordId, fields) {
  if (!['add', 'edit', 'delete'].includes(action)) throw new Error('عملية غير مدعومة.');
  if (!fields || typeof fields !== 'object' || Array.isArray(fields) || Object.keys(fields).some(key => !['name', 'duration', 'extra', 'days'].includes(key))) throw new Error('حقول غير مسموحة. المعادلات والنتائج محمية.');
  fields = { ...fields };
  const newModel = clone(model), found = [];
  for (const sheet of newModel.sheets) for (const group of sheet.groups) if (group.id === groupId) found.push([sheet, group]);
  if (found.length !== 1) throw new Error('مجموعة المحتوى غير موجودة.');
  const [sheet, group] = found[0];
  let record = group.records.find(item => item.id === recordId && !item.deleted);
  if (action !== 'add' && !record) throw new Error('السجل غير موجود أو سبق حذفه.');
  if (action === 'delete') {
    record.deletedCells = clone(rowCells(sheet, record.row));
    for (const [reference, cell] of Object.entries(sheet.cells)) if (rowOf(reference) === record.row) Object.assign(cell, { value: null, cached: null, formula: null, dirty: true });
    record.deleted = true; return newModel;
  }
  if (Object.hasOwn(fields, 'name')) {
    if (typeof fields.name !== 'string' || fields.name.length > 2000) throw new Error('اسم المحتوى يجب أن يكون نصًا لا يتجاوز 2000 حرف.');
    if (action === 'add' && !fields.name.trim()) throw new Error('يرجى إدخال اسم المحتوى الجديد.');
  } else if (action === 'add') throw new Error('يرجى إدخال اسم المحتوى الجديد.');
  if (Object.hasOwn(fields, 'duration')) {
    const old = record ? sheet.cells[inputRefs(sheet, record.row).duration]?.value ?? null : null;
    fields.duration = validateNumber(fields.duration, 'مدة المحتوى', { optional: action === 'edit' && old === null, positive: true });
  } else if (action === 'add') throw new Error('يرجى إدخال مدة المحتوى الجديد.');
  if (Object.hasOwn(fields, 'days')) {
    const old = record ? sheet.cells['H' + record.row]?.value ?? null : null;
    fields.days = validateNumber(fields.days, 'عدد الأيام', { optional: action === 'edit' && old === null, integer: true });
  }
  if (Object.hasOwn(fields, 'extra')) fields.extra = validateNumber(fields.extra, 'قيمة العمود C', { optional: true });
  if (action === 'add') {
    const candidate = item => {
      const cells = item.deleted ? item.deletedCells || {} : rowCells(sheet, item.row);
      return { score: Object.values(METRIC_COLS).filter(column => cells[column]?.formula).length, row: item.row, cells };
    };
    let candidates = group.records.map(candidate);
    if (!candidates.length && sheet.name === 'Sheet1') candidates = [{ score: 0, row: group.startRow, cells: rowCells(sheet, group.startRow) }];
    if (!candidates.length && sheet.name === 'Sheet2') candidates = sheet.groups.flatMap(otherGroup => otherGroup.records.map(candidate)).filter(item => item.score);
    if (!candidates.length) throw new Error('لا يحتوي الملف على قالب معادلات لإضافة سجل جديد إلى هذه المجموعة. استعد نسخة احتياطية.');
    const template = candidates.reduce((best, item) => item.score > best.score ? item : best);
    const row = insertRow(newModel, sheet, group, template.row, template.cells);
    record = { id: sheet.name + '-new-' + globalThis.crypto.randomUUID().replaceAll('-', ''), row, deleted: false }; group.records.push(record);
    if (sheet.name === 'Sheet1') fields = { extra: null, ...fields };
  }
  const refs = inputRefs(sheet, record.row);
  for (const [key, value] of Object.entries(fields)) {
    const reference = refs[key];
    if (reference === null || reference === undefined) { if (value !== null) throw new Error('الحقل ' + key + ' غير موجود في هذه الورقة.'); continue; }
    let cell = sheet.cells[reference];
    if (!cell) { cell = { value: null, cached: null, formula: null, attrs: { r: reference }, xml: null, originalRef: null, dirty: true }; sheet.cells[reference] = cell; }
    if (cell.formula) throw new Error('الخلية ' + reference + ' محسوبة ومحمية من تعديل الإدخال.');
    Object.assign(cell, { value, cached: value, dirty: true });
  }
  return newModel;
}

function serialize(root) { return strToU8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + serializer.serializeToString(root)); }
function appendValue(element, name, value) { const node = element.ownerDocument.createElementNS(MAIN, name); node.appendChild(element.ownerDocument.createTextNode(String(value))); element.appendChild(node); return node; }
function cellElement(reference, cell, calculator, name, doc) {
  const element = cell.xml ? doc.importNode(xml(cell.xml), true) : doc.createElementNS(MAIN, 'c');
  element.setAttribute('r', reference);
  for (const node of Array.from(element.childNodes)) if (node.namespaceURI === MAIN && ['f', 'v', 'is'].includes(node.localName)) element.removeChild(node);
  if (cell.formula) {
    element.removeAttribute('t'); appendValue(element, 'f', cell.formula.replace(/^=+/, ''));
    const [value, error] = calculator.safe(name, reference);
    if (error) { element.setAttribute('t', 'e'); appendValue(element, 'v', error); }
    else if (value !== null) appendValue(element, 'v', value);
    return element;
  }
  const value = cell.value ?? null;
  if (value === null) element.removeAttribute('t');
  else if (typeof value === 'string') {
    element.setAttribute('t', 'inlineStr'); const inline = doc.createElementNS(MAIN, 'is'), text = doc.createElementNS(MAIN, 't');
    text.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve'); text.appendChild(doc.createTextNode(value)); inline.appendChild(text); element.appendChild(inline);
  } else if (typeof value === 'boolean') { element.setAttribute('t', 'b'); appendValue(element, 'v', value ? '1' : '0'); }
  else { element.removeAttribute('t'); appendValue(element, 'v', value); }
  return element;
}

export async function exportXlsx(model, sourceBytes) {
  if (sourceBytes instanceof ArrayBuffer) sourceBytes = new Uint8Array(sourceBytes);
  if (await sha256(sourceBytes) !== model.sourceHash) throw new Error('ملف المصدر لا يطابق قاعدة البيانات الحالية.');
  const files = archive(sourceBytes), calculator = new Calculator(model), replacements = {};
  for (const sheet of model.sheets) {
    if (!Object.hasOwn(files, sheet.part)) throw new Error('رابط ورقة Excel غير صالح.');
    const root = xml(files[sheet.part]), oldData = child(root, 'sheetData');
    if (!oldData) throw new Error('بيانات ورقة Excel مفقودة.');
    const nextSibling = oldData.nextSibling; root.removeChild(oldData);
    for (const element of [root, ...Array.from(root.getElementsByTagName('*'))]) for (const key of ['ref', 'sqref', 'activeCell', 'topLeftCell']) if (element.hasAttribute(key)) {
      let reference = element.getAttribute(key);
      for (const at of sheet.insertions || []) reference = shiftRef(reference, at);
      element.setAttribute(key, reference);
    }
    const doc = root.ownerDocument, data = doc.createElementNS(MAIN, 'sheetData'), byRow = new Map();
    for (const reference of Object.keys(sheet.cells).sort(refSort)) { const row = rowOf(reference); if (!byRow.has(row)) byRow.set(row, []); byRow.get(row).push([reference, sheet.cells[reference]]); }
    const rows = [...new Set([...Object.keys(sheet.rowAttrs).map(Number), ...byRow.keys()])].sort((a, b) => a - b);
    for (const row of rows) {
      const rowElement = doc.createElementNS(MAIN, 'row');
      for (const [key, value] of Object.entries({ ...(sheet.rowAttrs[String(row)] || {}), r: String(row) })) rowElement.setAttribute(key, value);
      for (const [reference, cell] of byRow.get(row) || []) rowElement.appendChild(cellElement(reference, cell, calculator, sheet.name, doc));
      data.appendChild(rowElement);
    }
    root.insertBefore(data, nextSibling);
    const dimension = child(root, 'dimension');
    if (dimension) dimension.setAttribute('ref', attr(dimension, 'ref', 'A1').split(':')[0] + ':' + colLabel(sheet.maxColumn) + sheet.maxRow);
    replacements[sheet.part] = serialize(root);
  }
  const workbook = xml(files['xl/workbook.xml']);
  const definedNames = child(workbook, 'definedNames');
  if (definedNames) for (const item of children(definedNames, 'definedName')) {
    let text = item.textContent || '';
    for (const sheet of model.sheets) for (const at of sheet.insertions || []) { try { text = shiftFormula(text, '', sheet.name, at); } catch { /* Row-only print titles stay unchanged. */ } }
    item.textContent = text;
  }
  let calc = child(workbook, 'calcPr');
  if (!calc) { calc = workbook.ownerDocument.createElementNS(MAIN, 'calcPr'); workbook.appendChild(calc); }
  for (const [key, value] of Object.entries({ calcMode: 'auto', fullCalcOnLoad: '1', forceFullCalc: '1' })) calc.setAttribute(key, value);
  replacements['xl/workbook.xml'] = serialize(workbook);
  for (const part of ['xl/_rels/workbook.xml.rels', '[Content_Types].xml']) {
    const root = xml(files[part]);
    for (const node of Array.from(root.childNodes)) if (node.nodeType === 1 && [attr(node, 'Type', ''), attr(node, 'Target', ''), attr(node, 'PartName', '')].join('').includes('calcChain')) root.removeChild(node);
    replacements[part] = serialize(root);
  }
  const output = { ...files, ...replacements }; delete output['xl/calcChain.xml'];
  return zipSync(output, { level: 6 });
}

export const parse_xlsx = parseXlsx;
export const export_xlsx = exportXlsx;
export const baseline_mismatches = baselineMismatches;
