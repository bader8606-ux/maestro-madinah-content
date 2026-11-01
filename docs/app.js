'use strict';

const $ = (selector, root = document) => root.querySelector(selector);
const escapeHTML = value => String(value ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const nf = new Intl.NumberFormat('en-US', {maximumFractionDigits: 2});
const pf = new Intl.NumberFormat('en-US', {maximumFractionDigits: 2, style:'percent'});
const df = new Intl.DateTimeFormat('ar-SA-u-ca-gregory', {dateStyle:'medium', timeStyle:'short', timeZone:'Asia/Riyadh'});
let state = null;
let activeView = 'Sheet2';
let activeGroups = {};
let filters = {search:'', sort:'source', missing:false};
let sourceSheet = 'Sheet2';
let formulasOnly = false;
let dirty = false;
let formContext = null;
let previewSequence = 0;
let previewTimer;
let toastTimer;

function format(value, percent = false) {
  if (value === null || value === undefined || value === '') return '—';
  if (typeof value !== 'number') return escapeHTML(value);
  return Number.isFinite(value) ? (percent ? pf.format(value) : nf.format(value)) : 'غير متاح';
}
function date(value) { try { return df.format(new Date(value)); } catch { return '—'; } }
function revision() { return state?.document?.revision ?? state?.revision ?? 0; }
function allGroups(data = state) { return (data?.sheets || []).flatMap(s => s.groups || []); }
function currentSheet() { return state?.sheets?.find(s => s.name === activeView); }
function currentGroup() { const sheet = currentSheet(); return sheet?.groups?.find(g => g.id === activeGroups[activeView]) || sheet?.groups?.[0]; }
function getRecord(groupId, recordId, data = state) { return allGroups(data).find(g => g.id === groupId)?.records?.find(r => r.id === recordId); }
function sourceCode(record) { return `${activeView === 'Sheet1' ? 'A' : 'C'}${record.sourceRow}`; }
function recordMissing(record, sheet = activeView) {
  return !String(record.name ?? '').trim() || record.duration === null || record.duration === undefined ||
    (sheet === 'Sheet2' && (record.days == null || (record.missingOutputs || []).length > 0 || Object.values(record.errors || {}).some(Boolean)));
}
function metrics(record) { return record.metrics || {}; }
function setStatus(kind, text) { const node = $('#saveStatus'); node.className = `status-pill ${kind || ''}`; node.innerHTML = `<span class="status-dot"></span><span>${escapeHTML(text)}</span>`; }
function showSaved() { setStatus('', state?.savedAt ? `محفوظ في مساحة العمل · ${date(state.savedAt)}` : 'لا يوجد ملف محفوظ بعد'); }
function toast(message, error = false) {
  clearTimeout(toastTimer); const node = $('#toast'); node.textContent = message; node.className = `toast${error ? ' error' : ''}`;
  toastTimer = setTimeout(() => node.classList.add('hidden'), 5500);
}
async function api(path, options = {}) {
  return window.MaestroCloud.request(path,options);
}
function post(path, payload) { return api(path, {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify(payload)}); }
function updateState(data) {
  if (!window.MaestroAuth.getSession()) return;
  state = data;
  for (const sheet of state.sheets || []) {
    if (!(sheet.groups || []).some(g => g.id === activeGroups[sheet.name])) activeGroups[sheet.name] = sheet.groups?.[0]?.id;
  }
  $('#undoBtn').disabled = !state.canEdit || !state.canUndo;
  $('#importBtn').disabled = !state.canEdit;
  $('#saveBtn').disabled = !state.canEdit;
  $('#sourceInfo').textContent = `${state.document?.sourceName || 'Madinah content updated.xlsx'} · إصدار البيانات ${revision()}`;
  showSaved(); render();
}

function render() {
  if (!state) return;
  document.querySelectorAll('.nav-item').forEach(b => {
    const selected = b.dataset.view === activeView;
    b.classList.toggle('active', selected);
    b.setAttribute('aria-current', selected ? 'page' : 'false');
  });
  if (activeView === 'source') renderSource();
  else if (activeView === 'history') renderHistory();
  else if (activeView === 'help') renderHelp();
  else renderRecords();
}

function heading(title, subtitle, action = '') {
  return `<div class="page-heading"><div><div class="eyebrow">MAESTRO / MADINAH</div><h1 class="page-title">${escapeHTML(title)}</h1><p class="page-subtitle">${escapeHTML(subtitle)}</p></div>${action ? `<div class="heading-actions">${action}</div>` : ''}</div>`;
}
function warningsHTML(groupId = null) {
  const items = (state.warnings || []).filter(w => !groupId || w.groupId === groupId || (!w.groupId && w.sheet === activeView));
  if (!items.length) return '';
  return `<details class="notice warning"><summary>${items.length} ملاحظة من المرجع — اضغط للمراجعة</summary><ul class="warning-list">${items.map(w => `<li class="warning-item"><span>${escapeHTML(w.message)}</span>${w.ref ? `<code class="source-formula">${escapeHTML(w.sheet || '')}!${escapeHTML(w.ref)}</code>` : ''}</li>`).join('')}</ul></details>`;
}
function renderRecords() {
  const sheet = currentSheet(); const group = currentGroup();
  if (!sheet || !group) { $('#main').innerHTML = heading('مساحة العمل جاهزة',state.canEdit ? 'استورد ملف Excel المرجعي لبدء إدارة المحتوى. ستظهر معاينة قبل حفظه في المساحة المشتركة.' : 'لم يُضف ملف إلى مساحة العمل بعد. اطلب من مسؤول المساحة استيراده.') + '<div class="empty-state">لا يوجد ملف محفوظ بعد.</div>'; return; }
  const isCalc = activeView === 'Sheet2';
  const count = group.records?.length || 0;
  const missing = group.records.filter(r => recordMissing(r)).length;
  const title = isCalc ? 'حسابات العرض' : 'مدد المحتوى';
  const subtitle = isCalc ? 'تحرير المدخلات ومتابعة نتائج المعادلات الأصلية لكل مجموعة.' : 'المدد والأسماء والقيم الإضافية كما وردت في الورقة الأولى، دون ربط افتراضي بالورقة الثانية.';
  const total = group.totals || {};
  const durationSum = group.records.reduce((sum, r) => sum + (typeof r.duration === 'number' ? r.duration : 0), 0);
  const bar = durationSum > 0 ? `<div class="duration-bar" role="img" aria-label="توزيع مدد المقاطع داخل المجموعة">${group.records.filter(r => typeof r.duration === 'number' && r.duration > 0).map((r, i) => `<span class="duration-segment" style="width:${r.duration / durationSum * 100}%;background:${['#233389','#93268e','#4499d3'][i % 3]}" title="${escapeHTML(r.name || 'اسم غير مذكور')} · ${format(r.duration)}"></span>`).join('')}</div>` : '';
  $('#main').innerHTML = heading(title, subtitle, '<button class="btn primary" data-action="add">إضافة محتوى</button>') +
    `<div class="group-tabs" role="tablist" aria-label="مجموعات ${escapeHTML(sheet.name)}">${sheet.groups.map(g => `<button class="group-tab ${g.id === group.id ? 'active' : ''}" role="tab" aria-selected="${g.id === group.id}" data-group="${escapeHTML(g.id)}">${escapeHTML(g.title)}<span class="badge neutral">${g.records.length}</span></button>`).join('')}</div>` +
    `<div class="metrics-strip"><div class="metric"><span class="metric-label">سجلات المحتوى</span><strong class="metric-value">${count}</strong><span class="metric-note">${missing ? `${missing} سجلات بها بيانات ناقصة` : 'ضمن المجموعة المختارة'}</span></div><div class="metric"><span class="metric-label">إجمالي المدد${isCalc ? ' بالثانية' : ' (A)'}</span><strong class="metric-value">${format(total.duration)}</strong><span class="metric-note">${escapeHTML(group.totalRefs?.duration || '')} · نتيجة SUM الأصلية</span></div><div class="metric"><span class="metric-label">${isCalc ? 'إجمالي الساعات اليومية (G)' : 'الحقول الإضافية (C)'}</span><strong class="metric-value">${isCalc ? format(total.dailyHours) : group.records.filter(r => r.extra != null).length}</strong><span class="metric-note">${isCalc ? `${escapeHTML(group.totalRefs?.dailyHours || '')} · حسب معادلة المصدر` : 'دلالة العمود غير معنونة في المصدر'}</span></div></div>` +
    warningsHTML(group.id) +
    `<section class="panel"><div class="panel-heading"><div><h2 class="panel-title">${escapeHTML(group.title)}</h2><p class="panel-meta"><code>${escapeHTML(sheet.name)}!${escapeHTML(group.sourceRange || '')}</code> · ${isCalc ? 'النتائج المظللة محسوبة ومحمية' : 'الأسماء والمدد قابلة للتحرير'}</p></div><span class="badge neutral" id="resultCount"></span></div>${bar}<p class="section-note">توزيع مدد المحتوى ضمن هذه المجموعة، وليس تقديرًا مستقلًا لوقت البث.</p><div class="toolbar"><div class="search-field"><label class="sr-only" for="search">البحث في المحتوى</label><input class="field-input" id="search" type="search" placeholder="ابحث باسم المحتوى أو مرجعه…" value="${escapeHTML(filters.search)}"></div><label class="field"><span class="sr-only">ترتيب السجلات</span><select class="select-input" id="sort" aria-label="ترتيب السجلات"><option value="source">حسب ترتيب المصدر</option><option value="name">الاسم: أ–ي</option><option value="durationAsc">المدة: من الأقل</option><option value="durationDesc">المدة: من الأعلى</option>${isCalc ? '<option value="hoursDesc">الساعات اليومية: من الأعلى</option>' : ''}</select></label><label class="filter-check"><input id="missingOnly" type="checkbox" ${filters.missing ? 'checked' : ''}> البيانات الناقصة فقط</label><button class="btn secondary" data-action="clearFilters">مسح التصفية</button></div><div class="table-scroll"><table class="data-table" id="recordsTable"><thead><tr><th>المحتوى</th><th>${isCalc ? 'مدة الظهور (ثانية)' : 'المدة (A)'}</th>${isCalc ? '<th class="calculated">النسبة (D)</th><th class="calculated">زمن العرض (ثانية)</th><th class="calculated">عدد الظهور</th><th class="calculated">الساعات اليومية</th><th>عدد الأيام (H)</th><th class="calculated">الساعات الشهرية</th>' : '<th>الحقل الإضافي (C)</th>'}<th>المرجع</th><th>إجراءات</th></tr></thead><tbody id="recordsBody"></tbody></table></div><div class="footer-note">${isCalc ? 'المراجع والمعادلات مطابقة للمرجع، بما فيها ارتباط بعض المجموعات بإجمالي المجموعة الأولى. الخلايا الناقصة تبقى ناقصة ولا تُستبدل بصفر.' : 'حُفظت القيم كما وردت. لا تُحوّل القيم مثل 3.26 إلى دقائق أو ثوانٍ افتراضيًا، ولا يُفترض معنى إضافي للعمود C.'}</div></section>`;
  $('#sort').value = filters.sort;
  renderTable();
}
function renderTable() {
  const group = currentGroup(); if (!group || !$('#recordsBody')) return;
  const isCalc = activeView === 'Sheet2';
  const query = filters.search.trim().toLocaleLowerCase('ar');
  const records = group.records.filter(r => {
    const text = `${r.name || ''} ${r.duration ?? ''} ${r.days ?? ''} ${r.sourceRow || ''} ${sourceCode(r)}`.toLocaleLowerCase('ar');
    return (!query || text.includes(query)) && (!filters.missing || recordMissing(r));
  }).slice();
  const num = value => typeof value === 'number' ? value : -Infinity;
  if (filters.sort === 'name') records.sort((a, b) => (a.name || '').localeCompare(b.name || '', 'ar'));
  if (filters.sort === 'durationAsc') records.sort((a, b) => num(a.duration) - num(b.duration));
  if (filters.sort === 'durationDesc') records.sort((a, b) => num(b.duration) - num(a.duration));
  if (filters.sort === 'hoursDesc') records.sort((a, b) => num(metrics(b).dailyHours) - num(metrics(a).dailyHours));
  $('#resultCount').textContent = `${records.length} من ${group.records.length} سجل`;
  $('#recordsBody').innerHTML = records.length ? records.map(r => `<tr data-record="${escapeHTML(r.id)}"><td><div class="cell-name">${escapeHTML(r.name || 'اسم غير مذكور في المرجع')}</div>${recordMissing(r) ? '<span class="badge warn">بيانات ناقصة</span>' : ''}</td><td class="numeric" data-field="duration">${format(r.duration)}</td>${isCalc ? metricCell(r,'ratio',true)+metricCell(r,'seconds')+metricCell(r,'count')+metricCell(r,'dailyHours')+`<td class="numeric" data-field="days">${format(r.days)}</td>`+metricCell(r,'monthlyHours') : `<td class="numeric">${format(r.extra)}</td>`}<td><code class="cell-sub">${escapeHTML(sourceCode(r))}</code></td><td><div class="row-actions"><button class="btn icon" data-action="edit" data-id="${escapeHTML(r.id)}" aria-label="تعديل ${escapeHTML(r.name || `السجل ${r.sourceRow}`)}">تعديل</button><button class="btn icon danger" data-action="delete" data-id="${escapeHTML(r.id)}" aria-label="حذف ${escapeHTML(r.name || `السجل ${r.sourceRow}`)}">حذف</button></div></td></tr>`).join('') : `<tr><td colspan="${isCalc ? 10 : 5}"><div class="empty-state">لا توجد نتائج مطابقة. جرّب تغيير البحث أو التصفية.</div></td></tr>`;
  applyPermissions();
}
function metricCell(record, key, percent = false) {
  const formula = record.formulas?.[key]; const error = record.errors?.[key]; const value = metrics(record)[key];
  const title = error ? String(error) : formula ? `=${String(formula).replace(/^=/,'')}` : 'خلية خالية في المرجع؛ لم تُستكمل افتراضيًا';
  return `<td class="numeric calculated" data-field="${key}" title="${escapeHTML(title)}">${error ? `<span class="badge warn">${escapeHTML(error)}</span>` : format(value,percent)}</td>`;
}

function renderSource() {
  const sheet = state.sheets.find(s => s.name === sourceSheet) || state.sheets[0];
  if (!sheet) return;
  const rows = sheet.rows || [];
  const columns = sheet.name === 'Sheet1' ? ['A','B','C'] : ['B','C','D','E','F','G','H','I'];
  const visible = rows.filter(row => !formulasOnly || row.cells?.some(c => c.formula));
  $('#main').innerHTML = heading('المصدر والمعادلات', 'عرض مرجعي لكل الخلايا ذات المحتوى. تحرير المدخلات من صفحات إدارة المحتوى فقط.', '<a class="btn secondary" href="/api/original" download>تنزيل المرجع الأصلي</a>') +
    `<div class="metrics-strip"><div class="metric"><span class="metric-label">أوراق المرجع</span><strong class="metric-value">${state.sheets.length}</strong></div><div class="metric"><span class="metric-label">الخلايا المحفوظة</span><strong class="metric-value">${format(state.cellCount)}</strong></div><div class="metric"><span class="metric-label">الخلايا ذات المعادلات</span><strong class="metric-value">${format(state.formulaCount)}</strong></div></div>${warningsHTML()}<section class="panel"><div class="panel-heading"><h2 class="panel-title">الخلايا كما وردت في Excel</h2><label class="filter-check"><input id="formulasOnly" type="checkbox" ${formulasOnly ? 'checked' : ''}> المعادلات فقط</label></div><div class="tab-bar" role="tablist">${state.sheets.map(s => `<button class="tab-button ${s.name === sheet.name ? 'active' : ''}" role="tab" aria-selected="${s.name === sheet.name}" data-source-sheet="${escapeHTML(s.name)}">${escapeHTML(s.name)}</button>`).join('')}</div><div class="table-scroll"><table class="data-table source-table"><thead><tr><th>الصف</th>${columns.map(c => `<th>${c}</th>`).join('')}</tr></thead><tbody>${visible.map(row => `<tr><th class="numeric">${row.row}</th>${columns.map(col => {
      const cell = (row.cells || []).find(c => c.ref === `${col}${row.row}`);
      if (!cell) return '<td></td>';
      return `<td class="${cell.formula ? 'calculated' : ''}" data-cell="${escapeHTML(sheet.name)}!${escapeHTML(cell.ref)}"><div>${cell.error ? escapeHTML(cell.error) : format(cell.value)}</div>${cell.formula ? `<code class="source-formula">=${escapeHTML(String(cell.formula).replace(/^=/,''))}</code>` : ''}</td>`;
    }).join('')}</tr>`).join('')}</tbody></table></div></section>`;
}

function historyText(event) {
  const labels = {add:'إضافة محتوى',edit:'تعديل محتوى',delete:'حذف محتوى',undo:'تراجع عن تغيير',import:'استيراد ملف',save:'حفظ البيانات',restore:'استعادة نسخة احتياطية',initialize:'تحميل المرجع'};
  return labels[event.action] || event.action || 'تحديث البيانات';
}
function renderHistory() {
  const events = state.history || [];
  $('#main').innerHTML = heading('سجل التعديلات', 'تُسجل العمليات بعد حفظها فعليًا، مع إمكانية التراجع عن آخر تغيير.', '<button class="btn secondary" data-action="undo" '+($('#undoBtn').disabled ? 'disabled' : '')+'>تراجع عن آخر تغيير</button>') +
    `<section class="panel"><div class="panel-heading"><h2 class="panel-title">${events.length} عملية محفوظة</h2><a class="btn secondary" href="/api/backup" download>تنزيل نسخة احتياطية</a></div><div class="history-list">${events.length ? events.map(event => `<article class="history-item"><div><strong class="history-action">${escapeHTML(historyText(event))}</strong><p>${escapeHTML(typeof event.detail === 'object' ? JSON.stringify(event.detail) : event.detail || '')}</p></div><time class="history-date">${escapeHTML(date(event.at || event.timestamp))}</time></article>`).join('') : '<div class="empty-state">لم تُجرَ تعديلات بعد. المرجع الأصلي محفوظ دون تغيير.</div>'}</div></section>`;
}
function renderHelp() {
  $('#main').innerHTML = heading('الحفظ وطريقة الاستخدام', 'مساحة عمل مشتركة، مع تسجيل دخول وحفظ كل تغيير بعد تأكيده من قاعدة البيانات.') +
    `<div class="help-grid"><section class="panel"><div class="panel-heading"><h2 class="panel-title">التحرير والمزامنة</h2></div><div class="help-copy"><p>اختر الورقة والمجموعة، ثم أضف المحتوى أو عدّله. معاينة الحسابات فورية؛ يُعتمد التغيير عند الضغط على «حفظ التعديل» فقط.</p><p>تُحفظ البيانات في مساحة Supabase الخاصة ويستعيدها المستخدمون المخوّلون من أجهزتهم. تصل التحديثات تلقائيًا خلال نحو 15 ثانية، ويمكن طلبها بزر «تحديث البيانات».</p><p>إذا حفظ جهاز آخر أثناء تحريرك، يُرفض الحفظ القديم لمنع الكتابة فوق التعديل الأحدث. ألغِ النموذج وحدّث البيانات ثم أعد تعديلك. يلزم اتصال بالإنترنت للحفظ؛ لا يوجد حفظ صامت عند انقطاعه.</p></div></section><section class="panel"><div class="panel-heading"><h2 class="panel-title">الاستيراد والنسخ الاحتياطي</h2></div><div class="help-copy"><p>استورد ملف XLSX بنفس البنية أو نسخة JSON صادرة من النسخة السحابية. تظهر الخلايا المتغيرة قبل تأكيد الاستبدال. يمكن التراجع عن آخر عملية؛ ويُسجَّل التراجع كتعديل جديد.</p><p>النسخة الاحتياطية تشمل المصدر والبيانات الحالية وملخص سجل التعديلات. استعادتها تستعيد البيانات، وتُضاف عملية جديدة إلى السجل المشترك؛ لا يُعاد كتابة تاريخ العمليات.</p><p>التصدير يحافظ على الورقتين والمعادلات وأجزاء الملف الأخرى. يُفرّغ صف السجل المحذوف لحماية المراجع، وتُدرج السجلات الجديدة قبل الإجمالي.</p></div></section><section class="panel"><div class="panel-heading"><h2 class="panel-title">منطق الملف المرجعي</h2></div><div class="help-copy"><p>الورقتان مستقلتان. لا تُربط الأسماء المتشابهة افتراضيًا ولا يُخمّن معنى العمود C في Sheet1. بقيت الخلايا الناقصة والمراجع المرتبطة بالمجموعة الأولى كما في المصدر.</p><p>يدعم المحرك المراجع وSUM والحسابات المستخدمة في المرجع. تُرفض الصيغ غير المدعومة برسالة واضحة. النتائج المحسوبة محمية من التعديل العرضي داخل الواجهة.</p></div></section><section class="panel"><div class="panel-heading"><h2 class="panel-title">الصلاحيات والمرجع</h2></div><div class="help-copy"><p>الدخول وحده لا يمنح وصولًا؛ يحدد المسؤول أعضاء المساحة وصلاحيات العرض أو التحرير. لا تتضمن ملفات الموقع العامة بيانات المصنف.</p><p>استُخدم شعار مايسترو وخط GE Dinar One وألوان الملفات السابقة وفق موافقتك. المرجع الحالي: ${escapeHTML(state.document?.sourceName || 'لم يُستورد ملف بعد')}.</p><p><a href="/api/original" download>تنزيل المرجع الأصلي</a></p></div></section></div>`;
}

function openDialog(title, body, footer = '', wide = false) {
  if (!window.MaestroAuth.getSession()) return;
  const modal = $('#modal');
  modal.className = `dialog${wide ? ' wide' : ''}`;
  modal.innerHTML = `<div class="dialog-header"><h2 id="dialogTitle">${escapeHTML(title)}</h2><button class="btn icon" data-action="closeDialog" aria-label="إغلاق">×</button></div><div class="dialog-body">${body}</div>${footer ? `<div class="dialog-footer">${footer}</div>` : ''}`;
  if (!modal.open) modal.showModal();
}
function closeDialog() { $('#modal').close(); dirty = false; formContext = null; clearTimeout(previewTimer); previewSequence++; showSaved(); }

function openRecord(action, recordId = null) {
  const group = currentGroup(); if (!group) return;
  const record = recordId ? getRecord(group.id, recordId) : null;
  const isCalc = activeView === 'Sheet2';
  const days = record ? record.days : group.records.find(r => r.days != null)?.days ?? null;
  formContext = {action, groupId:group.id, recordId, isCalc, initialIds:new Set(group.records.map(r => r.id))};
  dirty = false;
  const blank = value => value == null ? '' : escapeHTML(value);
  openDialog(action === 'add' ? 'إضافة محتوى' : 'تعديل المحتوى',
    `<p class="section-note">${escapeHTML(group.title)} · ${escapeHTML(group.sheet || activeView)}. النتائج المحسوبة غير قابلة للتحرير.</p><form id="recordForm" novalidate><div class="form-grid"><label class="field form-span"><span class="field-label">اسم المحتوى${action === 'add' ? ' *' : ''}</span><input class="field-input" name="name" id="recordName" value="${escapeHTML(record?.name || '')}" autocomplete="off" ${action === 'add' ? 'required' : ''}></label><label class="field"><span class="field-label">${isCalc ? 'مدة الظهور بالثانية' : 'المدة كما وردت في A'}${action === 'add' ? ' *' : ''}</span><input class="field-input numeric" name="duration" id="recordDuration" inputmode="decimal" type="number" step="any" value="${blank(record?.duration)}" ${action === 'add' ? 'required' : ''}></label><label class="field"><span class="field-label">${isCalc ? 'عدد الأيام (H)' : 'الحقل الإضافي (C)'}</span><input class="field-input numeric" name="${isCalc ? 'days' : 'extra'}" id="recordExtra" inputmode="decimal" type="number" step="${isCalc ? '1' : 'any'}" value="${blank(isCalc ? days : record?.extra)}"></label></div><p class="input-error hidden" id="recordError" role="alert"></p><div class="live-preview" id="livePreview" aria-live="polite">${action === 'add' ? 'أدخل المدخلات لمعاينة الحسابات قبل الحفظ.' : 'جارٍ حساب المعاينة من المعادلات الأصلية…'}</div><div class="dialog-footer"><button class="btn secondary" type="button" data-action="closeDialog">إلغاء التغييرات</button><button class="btn primary" id="recordSave" type="submit">${action === 'add' ? 'إضافة وحفظ' : 'حفظ التعديل'}</button></div></form>`);
  $('#recordName').focus();
  if (record) schedulePreview();
}
function formFields() {
  const values = new FormData($('#recordForm'));
  const number = key => { const raw = String(values.get(key) ?? '').trim(); return raw === '' ? null : Number(raw); };
  const fields = {name:String(values.get('name') ?? ''), duration:number('duration')};
  if (formContext.isCalc) fields.days = number('days'); else fields.extra = number('extra');
  if (formContext.action === 'add' && !fields.name.trim()) throw new Error('أدخل اسم المحتوى.');
  if (fields.duration !== null && (!Number.isFinite(fields.duration) || fields.duration <= 0)) throw new Error('أدخل مدة موجبة أكبر من صفر.');
  if (formContext.action === 'add' && fields.duration === null) throw new Error('أدخل مدة المحتوى.');
  if (formContext.isCalc && fields.days !== null && (!Number.isInteger(fields.days) || fields.days < 0)) throw new Error('عدد الأيام يجب أن يكون عددًا صحيحًا غير سالب.');
  if (!formContext.isCalc && fields.extra !== null && (!Number.isFinite(fields.extra) || fields.extra < 0)) throw new Error('الحقل الإضافي يجب أن يكون رقمًا غير سالب أو فارغًا.');
  return fields;
}
function formPayload() { return {action:formContext.action, groupId:formContext.groupId, recordId:formContext.recordId, fields:formFields(), revision:revision()}; }
function recordError(message = '') { const node = $('#recordError'); if (!node) return; node.textContent = message; node.classList.toggle('hidden', !message); }
function schedulePreview() {
  clearTimeout(previewTimer); const ticket = ++previewSequence;
  previewTimer = setTimeout(async () => {
    if (!formContext || !$('#recordForm')) return;
    let payload;
    try { payload = formPayload(); recordError(); } catch (error) { recordError(error.message); $('#livePreview').textContent = 'صحح المدخلات لإظهار المعاينة.'; return; }
    try {
      const preview = await post('/api/record-preview', payload);
      if (ticket !== previewSequence || !formContext) return;
      const group = allGroups(preview).find(g => g.id === formContext.groupId);
      const record = formContext.action === 'add' ? group?.records?.find(r => !formContext.initialIds.has(r.id)) : group?.records?.find(r => r.id === formContext.recordId);
      if (!group || !record) { $('#livePreview').textContent = 'ستظهر النتيجة بعد حفظ المدخلات.'; return; }
      const values = formContext.isCalc ? [['الساعات اليومية',metrics(record).dailyHours],['الساعات الشهرية',metrics(record).monthlyHours],['إجمالي مدد المجموعة',group.totals?.duration]] : [['إجمالي مدد المجموعة',group.totals?.duration],['سجلات المجموعة',group.records.length]];
      $('#livePreview').innerHTML = `<p>معاينة مباشرة — لم تُحفظ بعد</p><div class="preview-values">${values.map(([label,value]) => `<div class="preview-stat"><span>${label}</span><strong class="numeric">${format(value)}</strong></div>`).join('')}</div>`;
    } catch (error) { if (ticket === previewSequence && $('#livePreview')) { $('#livePreview').textContent = 'تعذر حساب المعاينة.'; recordError(error.message); } }
  }, 220);
}
async function saveRecord() {
  if (!formContext) return;
  let payload; try { payload = formPayload(); } catch (error) { recordError(error.message); return; }
  $('#recordSave').disabled = true; setStatus('pending','جارٍ حفظ التعديل…');
  try { const data = await post('/api/records', payload); closeDialog(); updateState(data); toast('حُفظ المحتوى وتحدثت الحسابات.'); }
  catch (error) { recordError(error.message); setStatus('error','لم يُحفظ التعديل'); if ($('#recordSave')) $('#recordSave').disabled = false; }
}
function confirmDelete(recordId) {
  const group = currentGroup(); const record = getRecord(group.id,recordId); if (!record) return;
  formContext = {action:'delete',groupId:group.id,recordId};
  openDialog('تأكيد حذف المحتوى', `<p>هل تريد حذف «${escapeHTML(record.name || 'السجل غير المسمى')}»؟</p><p class="section-note">يمكنك التراجع بعد الحذف. عند تصدير Excel يُفرّغ صف هذا السجل حفاظًا على المراجع وبقية البنية.</p><p class="input-error hidden" id="deleteError" role="alert"></p>`, '<button class="btn secondary" data-action="closeDialog">إلغاء</button><button class="btn danger" data-action="confirmDelete" id="confirmDelete">حذف السجل</button>');
}
async function deleteRecord() {
  const context = {...formContext}; $('#confirmDelete').disabled = true;
  try { const data = await post('/api/records',{action:'delete',groupId:context.groupId,recordId:context.recordId,revision:revision()}); closeDialog(); updateState(data); toast('حُذف المحتوى. يمكنك التراجع عن العملية.'); }
  catch (error) { $('#deleteError').textContent = error.message; $('#deleteError').classList.remove('hidden'); $('#confirmDelete').disabled = false; }
}
async function undo() {
  if ($('#undoBtn').disabled) return;
  setStatus('pending','جارٍ التراجع…');
  try { updateState(await post('/api/undo',{revision:revision()})); toast('تم التراجع واستعادة البيانات السابقة.'); }
  catch (error) { showSaved(); toast(error.message,true); }
}
async function save() {
  if (formContext && $('#recordForm')) { await saveRecord(); return; }
  setStatus('pending','جارٍ التأكد من الحفظ…');
  try { updateState(await post('/api/save',{revision:revision()})); toast('البيانات محفوظة في مساحة العمل المشتركة.'); }
  catch (error) { setStatus('error','تعذر التأكد من الحفظ'); toast(error.message,true); }
}
async function importFile(file) {
  if (!file) return;
  if (!/\.(xlsx|json)$/i.test(file.name)) { toast('اختر ملف XLSX أو نسخة احتياطية JSON من هذه الأداة.',true); return; }
  if (file.size > 10 * 1024 * 1024) { toast('حجم الملف أكبر من الحد المدعوم (10 ميغابايت).',true); return; }
  openDialog('معاينة الاستيراد','<p>جارٍ قراءة الملف ومقارنة الخلايا والمعادلات، دون استبدال بياناتك…</p>');
  try {
    const preview = await api('/api/import-preview',{method:'POST',headers:{'Content-Type':/\.json$/i.test(file.name) ? 'application/json' : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet','X-Filename':encodeURIComponent(file.name)},body:await file.arrayBuffer()});
    formContext = {action:'import',token:preview.token};
    const summary = preview.summary || {};
    openDialog('مراجعة استبدال البيانات', `<p class="section-note">${escapeHTML(preview.sourceName || file.name)}</p><div class="notice warning">استيراد الملف يستبدل المصنف الحالي بالكامل بعد موافقتك فقط. تبقى النسخة السابقة قابلة للتراجع.</div><div class="metrics-strip"><div class="metric"><span class="metric-label">خلايا جديدة</span><strong class="metric-value">${summary.added || 0}</strong></div><div class="metric"><span class="metric-label">خلايا محذوفة</span><strong class="metric-value">${summary.removed || 0}</strong></div><div class="metric"><span class="metric-label">خلايا معدّلة</span><strong class="metric-value">${summary.changed || 0}</strong></div></div><p class="section-note">معادلات متغيرة: ${summary.formulasChanged || 0}. الاختلافات محسوبة حسب إحداثيات الخلايا؛ إدراج صف قد يغيّر مراجع خلايا لاحقة.</p><div class="table-scroll"><table class="data-table diff-table"><thead><tr><th>الخلية</th><th>قبل الاستيراد</th><th>بعد الاستيراد</th></tr></thead><tbody>${(preview.changes || []).slice(0,100).map(change => `<tr><td><code>${escapeHTML(change.sheet || '')}!${escapeHTML(change.ref || '')}</code></td><td>${escapeHTML(typeof change.before === 'object' ? JSON.stringify(change.before) : change.before ?? '—')}</td><td>${escapeHTML(typeof change.after === 'object' ? JSON.stringify(change.after) : change.after ?? '—')}</td></tr>`).join('') || '<tr><td colspan="3">لا توجد اختلافات في الخلايا.</td></tr>'}</tbody></table></div>${(preview.changes || []).length > 100 ? '<p class="section-note">تظهر أول 100 خلية متغيرة؛ الملخص يشمل جميع التغييرات.</p>' : ''}<p class="input-error hidden" id="importError" role="alert"></p>`, '<button class="btn secondary" data-action="closeDialog">إلغاء الاستيراد</button><button class="btn primary" data-action="confirmImport" id="confirmImport">تأكيد استبدال البيانات</button>', true);
  } catch (error) { openDialog('تعذر استيراد الملف',`<p class="input-error" role="alert">${escapeHTML(error.message)}</p>`,'<button class="btn secondary" data-action="closeDialog">إغلاق</button>'); }
}
async function confirmImport() {
  $('#confirmImport').disabled = true;
  try { const data = await post('/api/import-confirm',{token:formContext.token,revision:revision()}); closeDialog(); activeGroups = {}; filters = {search:'',sort:'source',missing:false}; updateState(data); toast('حُفظ الملف المستورد بعد التأكيد.'); }
  catch (error) { $('#importError').textContent = error.message; $('#importError').classList.remove('hidden'); $('#confirmImport').disabled = false; }
}

document.addEventListener('click', async event => {
  const download=event.target.closest('[data-download],a[href^="/api/"]');
  if (download) {
    event.preventDefault();
    try { await window.MaestroCloud.download(download.dataset.download || download.getAttribute('href').split('/').pop()); }
    catch(error) { toast(error.message,true); }
    return;
  }
  const button = event.target.closest('button'); if (!button || button.disabled) return;
  if (button.dataset.view) { activeView = button.dataset.view; filters = {search:'',sort:'source',missing:false}; render(); return; }
  if (button.dataset.group) { activeGroups[activeView] = button.dataset.group; filters.search = ''; renderRecords(); return; }
  if (button.dataset.sourceSheet) { sourceSheet = button.dataset.sourceSheet; renderSource(); return; }
  const action = button.dataset.action;
  if (action === 'add') openRecord('add');
  if (action === 'edit') openRecord('edit',button.dataset.id);
  if (action === 'delete') confirmDelete(button.dataset.id);
  if (action === 'confirmDelete') await deleteRecord();
  if (action === 'closeDialog') closeDialog();
  if (action === 'undo') await undo();
  if (action === 'confirmImport') await confirmImport();
  if (action === 'clearFilters') { filters = {search:'',sort:'source',missing:false}; renderRecords(); }
});
document.addEventListener('input', event => {
  if (event.target.id === 'search') { filters.search = event.target.value; renderTable(); }
  if (event.target.closest('#recordForm')) { dirty = true; setStatus('pending','تعديلات في النموذج لم تُحفظ'); schedulePreview(); }
});
document.addEventListener('change', event => {
  if (event.target.id === 'sort') { filters.sort = event.target.value; renderTable(); }
  if (event.target.id === 'missingOnly') { filters.missing = event.target.checked; renderTable(); }
  if (event.target.id === 'formulasOnly') { formulasOnly = event.target.checked; renderSource(); }
});
document.addEventListener('submit', event => { if (event.target.id === 'recordForm') { event.preventDefault(); saveRecord(); } });
$('#saveBtn').addEventListener('click',save);
$('#refreshBtn').addEventListener('click',async()=>{ try { updateState(await api('/api/state')); toast('تم تحديث البيانات من مساحة العمل.'); } catch(error){ toast(error.message,true); } });
$('#undoBtn').addEventListener('click',undo);
$('#importBtn').addEventListener('click',() => { $('#importFile').value = ''; $('#importFile').click(); });
$('#importFile').addEventListener('change',event => importFile(event.target.files?.[0]));
$('#modal').addEventListener('cancel',() => { dirty = false; formContext = null; clearTimeout(previewTimer); previewSequence++; showSaved(); });
window.addEventListener('beforeunload',event => { if (dirty) { event.preventDefault(); event.returnValue = ''; } });
document.addEventListener('keydown',event => { if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 's') { event.preventDefault(); save(); } });

function applyPermissions() {
  if (!state || state.canEdit) return;
  document.querySelectorAll('[data-action="add"],[data-action="edit"],[data-action="delete"],[data-action="undo"]').forEach(button=>{button.disabled=true;button.title='صلاحية العرض فقط';});
}
async function openWorkspace() {
  try { updateState(await api('/api/state')); }
  catch(error) {
    setStatus('error','تعذر فتح مساحة العمل');
    $('#main').innerHTML = heading('تعذر فتح البيانات','تحقق من الاتصال وصلاحية حسابك، ثم أعد المحاولة.') + `<div class="notice warning">${escapeHTML(error.message)}</div><button class="btn primary" id="retryBtn">إعادة المحاولة</button>`;
    $('#retryBtn').addEventListener('click',openWorkspace);
  }
}
window.MaestroCloud.subscribe((next,error)=>{
  if (error) {
    if ([401,403].includes(error.status)) { state=null;window.MaestroCloud.clear();$('#modal').close();$('#modal').innerHTML='';$('#main').textContent=error.message; }
    setStatus('error','تعذرت المزامنة — تحقق من الاتصال والصلاحيات'); return;
  }
  if (!dirty && !$('#modal').open) { updateState(next); toast('وصل تحديث من مساحة العمل.'); }
});
window.MaestroAuth.subscribe((event)=>{
  if (event==='SIGNED_IN' || event==='INITIAL_SESSION' && window.MaestroAuth.getSession()) openWorkspace();
  if (event==='SIGNED_OUT') { state=null;window.MaestroCloud.clear();dirty=false;formContext=null;$('#modal').close();$('#modal').innerHTML='';$('#main').replaceChildren();$('#sourceInfo').textContent='';clearTimeout(previewTimer);previewSequence++; }
});
window.MaestroAuth.start({appSelector:'#app',beforeSignOut:()=> !dirty || window.confirm('توجد تغييرات لم تُحفظ. هل تريد تسجيل الخروج وإلغاء التغييرات؟')});
