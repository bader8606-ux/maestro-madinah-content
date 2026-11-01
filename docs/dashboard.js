/* Source-based Maestro dashboard. No workbook data is embedded in this file. */
(function (global) {
  'use strict';

  const instances = new Map();
  const colors = ['#233389', '#93268e', '#4499d3', '#5768ad', '#b35bac', '#76b5df'];
  const numberFormat = new Intl.NumberFormat('en-US', {numberingSystem:'latn', maximumFractionDigits:2});
  let serial = 0;
  const finite = value => typeof value === 'number' && Number.isFinite(value);
  const escape = value => String(value ?? '').replace(/[&<>"']/g, character => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[character]));
  const latin = value => String(value ?? '').replace(/[٠-٩۰-۹]/g, character => String(character.charCodeAt(0) - (character <= '٩' ? 1632 : 1776)));
  const text = value => escape(latin(value));
  const format = value => finite(value) ? numberFormat.format(value) : 'غير متوفر';
  const precise = value => finite(value) ? String(value) : 'غير متوفر';
  const nameOf = record => String(record.name ?? '').trim() || 'غير متوفر';
  const defaultFilters = state => {
    const sheet = state?.sheets?.find(item => item.name === 'Sheet2') || state?.sheets?.[0];
    return {sheet:sheet?.name || 'Sheet2',groupId:sheet?.groups?.[0]?.id || 'all',recordId:'all',search:'',sort:'source',direction:'asc',missing:false};
  };
  const fields = ['duration','count','dailyHours','monthlyHours','seconds'];
  const labels = {
    duration:'مجموع مدد المحتوى',count:'مرات الظهور اليومية',dailyHours:'ساعات العرض اليومية',
    monthlyHours:'ساعات العرض الشهرية',seconds:'إجمالي وقت العرض اليومي'
  };
  function unit(field, sheet) {
    if (field === 'duration') return sheet === 'Sheet2' ? 'ثانية' : 'وحدة غير محددة';
    return {count:'مرة / يوم',dailyHours:'ساعة / يوم',monthlyHours:'ساعة / شهر',seconds:'ثانية / يوم'}[field] || '';
  }
  function column(field, sheet) {
    return field === 'duration' ? (sheet === 'Sheet2' ? 'C' : 'A') : {count:'F',dailyHours:'G',monthlyHours:'I',seconds:'E'}[field];
  }
  function valueOf(record, field) {
    if (field === 'duration') return record.duration;
    return Object.hasOwn(record,field) ? record[field] : record.metrics?.[field];
  }
  function reference(record, field, sheet) {
    const ref = field === 'duration' ? record.inputRefs?.duration : column(field,sheet) + record.sourceRow;
    return sheet + '!' + (ref || column(field,sheet) + record.sourceRow);
  }
  function coverageHTML(coverage, applicable = true) {
    if (!applicable || coverage?.applicable === false) return '<span class="md-coverage">غير موجود في هذه الورقة</span>';
    const available = Number(coverage?.available || 0), total = Number(coverage?.total || 0);
    return `<span class="md-coverage ${coverage?.partial ? 'is-partial' : ''}"><span class="md-digits">${available} / ${total}</span> قيمة متاحة${coverage?.partial ? ' · مجموع جزئي' : ''}${coverage?.errors ? ` · ${Number(coverage.errors)} أخطاء` : ''}</span>`;
  }
  function description(data) {
    return [data.name,`القيمة: ${precise(data.value)} ${data.unit || ''}`,data.ref || '',finite(data.share) ? `النسبة: ${precise(data.share)}%` : '',finite(data.denominator) ? `المقام: ${precise(data.denominator)} ${data.denominatorUnit || data.unit || ''}` : '',data.note || ''].filter(Boolean).join('\n');
  }
  function tooltipAttrs(data) {
    return `data-dash-tooltip="${escape(JSON.stringify(data))}" title="${text(description(data))}"`;
  }
  function axisDomain(values) {
    const valid = values.filter(finite);
    const low = Math.min(0,...valid), high = Math.max(0,...valid);
    return {low,high,span:high-low || 1,zero:(-low/(high-low || 1))*100};
  }
  function percentAllowed(records, field) {
    const values = records.map(record => valueOf(record,field)).filter(finite);
    return values.length > 0 && values.every(value => value >= 0) && values.reduce((sum,value) => sum+value,0) > 0;
  }
  function metricHeader(title, eyebrow, note, action = '') {
    return `<header class="md-panel-heading"><div><span class="md-kicker">${text(eyebrow)}</span><h2>${text(title)}</h2>${note ? `<p>${text(note)}</p>` : ''}</div>${action}</header>`;
  }
  function selectionButton(record, inner, data, filters, className = '') {
    return `<button type="button" class="${className}" data-dash-record="${escape(record.id)}" aria-pressed="${filters.recordId === record.id}" ${tooltipAttrs(data)}>${inner}</button>`;
  }
  function horizontalChart(records, field, data, instance, options = {}) {
    if (field==='duration' && data.sheet==='Sheet1' && data.coverage?.duration?.aggregatable===false) return '<div class="md-chart-empty">اختر مجموعة واحدة لعرض مدد Sheet1. لا توضع وحدات غير مؤكدة عبر المجموعات على محور مشترك.</div>';
    const percentage = options.percentage === true;
    const valid = records.filter(record => finite(valueOf(record,field)));
    if (!valid.length) return '<div class="md-chart-empty">لا توجد قيم متاحة ضمن التصفية الحالية.</div>';
    const denominator = valid.reduce((sum,record) => sum + valueOf(record,field),0);
    const plotted = record => percentage ? valueOf(record,field)/denominator*100 : valueOf(record,field);
    const domain = axisDomain(valid.map(plotted));
    const currentUnit = percentage ? '%' : unit(field,data.sheet);
    return `<div class="md-horizontal-chart"><div class="md-chart-axis"><span>${format(domain.low)}</span><span>${format(domain.high)} ${text(currentUnit)}</span></div><div class="md-bar-list">${records.map((record,index) => {
      const original = valueOf(record,field), value = finite(original) ? plotted(record) : null;
      const width = finite(value) ? Math.abs(value)/domain.span*100 : 0;
      const start = finite(value) ? domain.zero + (value < 0 ? -width : 0) : domain.zero;
      const share = denominator > 0 && finite(original) && original >= 0 ? original/denominator*100 : null;
      const tip = {name:nameOf(record),value:original,unit:unit(field,data.sheet),share,denominator,ref:reference(record,field,data.sheet),note:'اضغط لتصفية جميع عناصر اللوحة بهذا المحتوى. اضغط مرة أخرى لإلغاء تحديده.'};
      const inner = `<span class="md-bar-label">${text(nameOf(record))}<code>${text(reference(record,field,data.sheet))}</code></span><span class="md-bar-track" aria-hidden="true"><span class="md-zero" style="inset-inline-start:${domain.zero}%"></span>${finite(value) ? `<span class="md-bar-fill" style="inset-inline-start:${Math.max(0,start)}%;width:${Math.min(100,width)}%;background:${colors[index%colors.length]}"></span>` : '<span class="md-bar-unavailable">غير متوفر</span>'}</span><span class="md-bar-value md-digits">${format(value)}${percentage && finite(value) ? '%' : ''}</span>`;
      return selectionButton(record,inner,tip,data.filters,'md-bar-row' + (data.filters.recordId === record.id ? ' is-selected' : ''));
    }).join('')}</div></div>`;
  }
  function summaryCards(data) {
    const countTip={name:'عدد سجلات المحتوى',value:data.records.length,unit:'سجل',note:'عدد الصفوف المعروضة بعد التصفية؛ تُحسب الأسماء المكررة سجلات مستقلة وفق مراجعها في المصدر.'};
    const countCard=`<article class="md-metric" tabindex="0" ${tooltipAttrs(countTip)}><span class="md-metric-label">عدد سجلات المحتوى<code>ROWS</code></span><strong class="md-metric-value md-digits">${format(data.records.length)}</strong><span class="md-metric-unit">سجل</span><span class="md-coverage">كل صف محتوى سجل مستقل · بعد التصفية</span></article>`;
    return `<div class="md-metric-grid">${countCard}${fields.map(field => {
      const applicable = data.sheet === 'Sheet2' || field === 'duration';
      const value = applicable ? data.totals?.[field] : null;
      const reason = data.provenance?.metrics?.[field]?.reason;
      const tip = {name:labels[field],value,unit:unit(field,data.sheet),ref:applicable ? `عمود ${column(field,data.sheet)} · مجموع القيم المتاحة بعد التصفية` : '',note:reason || (data.scope?.aggregateBasis === 'cumulative_across_groups' ? 'مجموع تراكمي عبر المجموعات المختارة، وليس وقتًا منقضيًا على شاشة واحدة.' : '')};
      return `<article class="md-metric ${data.coverage?.[field]?.partial ? 'is-partial' : ''}" tabindex="0" ${tooltipAttrs(tip)}><span class="md-metric-label">${text(labels[field])}<code>${applicable ? column(field,data.sheet) : '—'}</code></span><strong class="md-metric-value ${!finite(value) ? 'is-unavailable' : 'md-digits'}">${format(value)}</strong><span class="md-metric-unit">${applicable ? text(unit(field,data.sheet)) : 'غير متوفر في المصدر'}</span>${reason ? `<span class="md-coverage is-partial">لا تُجمع وحدات غير مؤكدة عبر المجموعات</span>` : coverageHTML(data.coverage?.[field],applicable)}</article>`;
    }).join('')}</div>`;
  }
  function shareChart(data) {
    const shares = data.shares || [];
    const available = shares.filter(item => finite(item.percent));
    const denominator = shares.find(item => finite(item.denominator))?.denominator;
    const heading = metricHeader('توزيع وقت العرض اليومي','حصة المحتوى · E','النسبة من مجموع E المتاح بعد جميع الفلاتر الحالية.');
    if (data.sheet !== 'Sheet2') return `<section class="md-panel">${heading}<div class="md-chart-empty">لا تتضمن Sheet1 زمن عرض محسوبًا في E. لا تُستنتج نسبة بديلة.</div></section>`;
    if (data.records.some(record=>finite(record.seconds) && record.seconds<0)) return `<section class="md-panel">${heading}<div class="md-chart-empty">لا يُعرض توزيع نسبي عند وجود قيم E سالبة. تبقى القيم الأصلية واضحة في الجدول.</div></section>`;
    if (!available.length) return `<section class="md-panel">${heading}<div class="md-chart-empty">النسبة غير متوفرة: لا يوجد مجموع E موجب متاح ضمن التصفية.</div>${coverageHTML(data.coverage?.seconds)}</section>`;
    return `<section class="md-panel md-share-panel">${heading}<div class="md-share-body"><div class="md-share-total"><span>المقام المستخدم</span><strong class="md-digits">${format(denominator)} <small>ثانية / يوم</small></strong></div><div class="md-share-rail" aria-label="توزيع وقت العرض اليومي على المحتوى">${available.map((item,index) => {
      const record = data.records.find(row => row.id === (item.recordId || item.id));
      const tip = {name:item.name || (record ? nameOf(record) : 'غير متوفر'),value:item.value,unit:'ثانية / يوم',share:item.percent,denominator:item.denominator,ref:data.sheet+'!'+(item.ref || 'E'+item.sourceRow)};
      return `<button type="button" tabindex="-1" data-dash-record="${escape(item.recordId || item.id)}" class="md-share-segment" style="width:${Math.max(0,item.percent)}%;background:${colors[index%colors.length]}" aria-label="${text(description(tip))}" aria-pressed="${data.filters.recordId === (item.recordId || item.id)}" ${tooltipAttrs(tip)}></button>`;
    }).join('')}</div>${coverageHTML(data.coverage?.seconds)}<div class="md-share-list">${shares.map((item,index) => {
      const tip = {name:item.name || 'غير متوفر',value:item.value,unit:'ثانية / يوم',share:item.percent,denominator:item.denominator,ref:data.sheet+'!'+(item.ref || 'E'+item.sourceRow)};
      return `<button type="button" class="md-share-item" data-dash-record="${escape(item.recordId || item.id)}" aria-pressed="${data.filters.recordId === (item.recordId || item.id)}" ${tooltipAttrs(tip)}><span class="md-swatch" style="background:${colors[index%colors.length]}"></span><span class="md-share-name">${text(item.name || 'غير متوفر')}<code>${text(tip.ref)}</code></span><span class="md-share-numbers"><strong class="md-digits">${format(item.percent)}${finite(item.percent) ? '%' : ''}</strong><small class="md-digits">${format(item.value)} ث / يوم</small></span></button>`;
    }).join('')}</div></div></section>`;
  }
  function appearanceChart(data, instance) {
    const field = 'count', records = data.records;
    const heading = metricHeader('مرات الظهور حسب المحتوى','العمود F · يومي','تُعرض القيم الأصلية بكسورها كما هي، دون تقريبها إلى عدد صحيح.');
    if (data.sheet !== 'Sheet2') return `<section class="md-panel md-appearance-panel">${heading}<div class="md-chart-empty">لا يحتوي المصدر في Sheet1 على عمود لمرات الظهور.</div></section>`;
    const valid = records.filter(record => finite(valueOf(record,field)));
    if (!valid.length) return `<section class="md-panel md-appearance-panel">${heading}<div class="md-chart-empty">لا توجد قيم F متاحة ضمن التصفية الحالية.</div></section>`;
    const percentage = instance.mode === 'percent' && percentAllowed(records,field);
    const denominator = valid.reduce((sum,record) => sum+valueOf(record,field),0);
    const plotted = record => percentage ? valueOf(record,field)/denominator*100 : valueOf(record,field);
    const domain = axisDomain(valid.map(plotted));
    const zeroY = 158 - domain.zero/100*148;
    return `<section class="md-panel md-appearance-panel">${heading}<div class="md-appearance-meta"><span>${percentage ? 'النسبة من مجموع F المتاح في التصفية' : 'الوحدة: مرة / يوم'} · ${format(domain.high)}${percentage ? '%' : ''} أقصى محور</span>${coverageHTML(data.coverage?.count)}</div><div class="md-columns-scroll" tabindex="0" aria-label="مخطط أعمدة مرات الظهور؛ اسحب أفقيًا لعرض جميع الأسماء"><div class="md-columns" style="--column-count:${records.length}">${records.map((record,index) => {
      const original = valueOf(record,field), value = finite(original) ? plotted(record) : null;
      const h = finite(value) ? Math.abs(value)/domain.span*148 : 0;
      const y = finite(value) ? (value >= 0 ? zeroY-h : zeroY) : zeroY;
      const tip = {name:nameOf(record),value:original,unit:'مرة / يوم',share:denominator>0 && finite(original) && original>=0 ? original/denominator*100 : null,denominator,ref:reference(record,field,data.sheet)};
      const inner = `<span class="md-column-value md-digits">${format(value)}${percentage && finite(value) ? '%' : ''}</span><svg viewBox="0 0 80 168" class="md-column-graphic" aria-hidden="true"><line x1="0" y1="${zeroY}" x2="80" y2="${zeroY}" stroke="#cfd7e7" stroke-width="1"/>${finite(value) ? `<rect x="20" y="${Math.max(3,y)}" width="40" height="${Math.max(value===0 ? 1 : 0,h)}" rx="3" fill="${colors[index%colors.length]}"/>` : '<text x="40" y="88" text-anchor="middle" fill="#8a93a5" font-size="20">—</text>'}</svg><span class="md-column-label">${text(nameOf(record))}</span><code>${text(reference(record,field,data.sheet))}</code>`;
      return selectionButton(record,inner,tip,data.filters,'md-column' + (data.filters.recordId === record.id ? ' is-selected' : ''));
    }).join('')}</div></div></section>`;
  }
  function comparisonChart(data, instance) {
    const field = instance.comparisonMetric;
    const options = fields.filter(item => data.sheet === 'Sheet2' || item === 'duration');
    const action = `<label class="md-compare-select"><span class="sr-only">مؤشر المقارنة بين المجموعات</span><select data-dash-filter="comparisonMetric" aria-label="مؤشر المقارنة بين المجموعات">${options.map(item => `<option value="${item}" ${field===item ? 'selected' : ''}>${text(labels[item])} · ${column(item,data.sheet)}</option>`).join('')}</select></label>`;
    const heading = metricHeader('مقارنة المجموعات','نطاق التصفية الحالي','اختر «جميع مجموعات الورقة» من الأعلى للمقارنة بينها. لا تُخلط الوحدات في محور واحد.',action);
    const groups = data.comparisons || [];
    const values = groups.map(group => group.selectedTotals?.[field]);
    if (!groups.length || !values.some(finite)) return `<section class="md-panel md-comparison-panel">${heading}<div class="md-chart-empty">لا توجد قيم متاحة لهذا المؤشر ضمن التصفية الحالية.</div></section>`;
    if (data.sheet==='Sheet1' && data.coverage?.duration?.aggregatable===false) return `<section class="md-panel md-comparison-panel">${heading}<p class="md-chart-hint">لكل مجموعة أساس غير مؤكد؛ لذلك تُعرض قيمها الأصلية منفصلة، دون محور مشترك أو مجموع أو نسب بينها.</p><div class="md-unknown-comparisons">${groups.map(group=>`<button type="button" data-dash-group="${escape(group.id)}"><span>${text(group.title)}<code>${text(data.sheet+'!'+group.sourceRange)}</code></span><span class="md-digits">${format(group.selectedTotals?.duration)}</span></button>`).join('')}</div></section>`;
    const domain = axisDomain(values);
    const chart = `<div class="md-horizontal-chart"><div class="md-chart-axis"><span>${format(domain.low)}</span><span>${format(domain.high)} ${text(unit(field,data.sheet))}</span></div><div class="md-bar-list">${groups.map((group,index) => {
      const value = group.selectedTotals?.[field];
      const width = finite(value) ? Math.abs(value)/domain.span*100 : 0;
      const start = finite(value) ? domain.zero + (value<0 ? -width : 0) : domain.zero;
      const tip = {name:group.title,value,unit:unit(field,data.sheet),ref:data.sheet+'!'+group.sourceRange,note:'مجموع القيم المتاحة للصفوف المطابقة. اضغط لاختيار هذه المجموعة.'};
      return `<button type="button" class="md-bar-row md-group-row ${data.filters.groupId===group.id ? 'is-selected' : ''}" data-dash-group="${escape(group.id)}" aria-pressed="${data.filters.groupId===group.id}" ${tooltipAttrs(tip)}><span class="md-bar-label">${text(group.title)}<code>${text(data.sheet+'!'+group.sourceRange)}</code>${coverageHTML(group.coverage?.[field])}</span><span class="md-bar-track" aria-hidden="true"><span class="md-zero" style="inset-inline-start:${domain.zero}%"></span>${finite(value) ? `<span class="md-bar-fill" style="inset-inline-start:${Math.max(0,start)}%;width:${Math.min(100,width)}%;background:${colors[index%colors.length]}"></span>` : ''}</span><span class="md-bar-value md-digits">${format(value)}</span></button>`;
    }).join('')}</div></div>`;
    const sourceFields = data.sheet==='Sheet2' ? ['duration','seconds','dailyHours'] : ['duration'];
    const originals = `<details class="md-originals"><summary>مراجعة إجماليات Excel الأصلية</summary><p>الجدول التالي يعرض خلايا الإجماليات كما وردت في المصدر. قد تختلف عن جمع الصفوف المصفاة أو عن مجموع عمود آخر.</p><div class="table-scroll"><table class="md-source-table"><thead><tr><th>المجموعة</th>${sourceFields.map(item => `<th>${text(labels[item])}<small>${text(unit(item,data.sheet))}</small></th>`).join('')}</tr></thead><tbody>${groups.map(group => `<tr><td>${text(group.title)}</td>${sourceFields.map(item => `<td><span class="md-digits">${format(group.originalTotals?.[item])}</span><code>${text(group.originalTotalRefs?.[item] || '')}</code></td>`).join('')}</tr>`).join('')}</tbody></table></div></details>`;
    return `<section class="md-panel md-comparison-panel">${heading}${chart}${originals}</section>`;
  }
  function linkedTable(data, instance) {
    const calc = data.sheet==='Sheet2';
    const headings = calc ? [['duration','مدة المحتوى'],['count','الظهور اليومي'],['seconds','وقت العرض اليومي'],['dailyHours','الساعات اليومية'],['monthlyHours','الساعات الشهرية']] : [['duration','المدة الأصلية']];
    return `<section class="md-panel md-linked-panel">${metricHeader('تفاصيل المحتوى','جدول مرتبط بكل الرسوم',`${data.records.length} سجل ضمن التصفية. اضغط اسم المحتوى لتحديده أو إلغاء تحديده.`)}<div class="table-scroll" tabindex="0" aria-label="تفاصيل المحتوى؛ جدول قابل للتمرير أفقيًا"><table class="md-data-table"><thead><tr><th>المحتوى والمجموعة</th>${headings.map(([field,label]) => `<th>${text(label)}<small>${column(field,data.sheet)} · ${text(unit(field,data.sheet))}</small></th>`).join('')}${calc ? '<th>عدد الأيام<small>H · من المصدر</small></th>' : '<th>الحقل الإضافي<small>C · غير معنْون في المصدر</small></th>'}<th>المرجع</th><th>إجراء</th></tr></thead><tbody>${data.records.map(record => `<tr class="${data.filters.recordId===record.id ? 'is-selected' : ''}"><td><button type="button" class="md-table-name" data-dash-record="${escape(record.id)}" aria-pressed="${data.filters.recordId===record.id}">${text(nameOf(record))}</button><span class="md-table-group">${text(record.groupTitle)}</span></td>${headings.map(([field]) => {
      const value = valueOf(record,field), error=record.errors?.[field];
      const tip={name:nameOf(record),value,unit:unit(field,data.sheet),ref:reference(record,field,data.sheet),note:error ? String(error) : !finite(value) ? 'خلية غير متاحة؛ لم تُستبدل بصفر.' : ''};
      return `<td class="${field==='duration' ? '' : 'md-calculated'}" tabindex="0" ${tooltipAttrs(tip)}><span class="${finite(value) ? 'md-digits' : 'md-cell-missing'}">${format(value)}</span>${error ? `<small class="md-cell-error">${text(error)}</small>` : ''}</td>`;
    }).join('')}<td><span class="${finite(calc ? record.days : record.extra) ? 'md-digits' : 'md-cell-missing'}">${format(calc ? record.days : record.extra)}</span></td><td><code>${text(data.sheet+'!'+(record.inputRefs?.name || 'B'+record.sourceRow))}</code></td><td><button type="button" class="btn icon md-edit" data-dash-edit="${escape(record.id)}" data-dash-edit-group="${escape(record.groupId)}" ${!instance.state.canEdit || typeof instance.options.onEdit!=='function' ? 'disabled title="صلاحية العرض فقط"' : ''} aria-label="تعديل ${text(nameOf(record))}">تعديل</button></td></tr>`).join('') || `<tr><td colspan="${calc ? 9 : 5}"><div class="md-chart-empty">لا توجد سجلات مطابقة. امسح الفلاتر أو غيّر البحث.</div></td></tr>`}</tbody></table></div><div class="md-table-footnote">الخلايا المظللة محسوبة من معادلات المصدر. القيم المفقودة لا تُستبدل بصفر؛ وتعرض البطاقات مجموع القيم المتاحة فقط.</div></section>`;
  }
  function warningHTML(data) {
    const warnings = data.warnings || [];
    if (!warnings.length) return '';
    return `<details class="md-warnings"><summary>${warnings.length} ملاحظة من ملف المصدر</summary><ul>${warnings.map(item => `<li><span>${text(item.message)}</span>${item.ref ? `<code>${text((item.sheet || data.sheet)+'!'+item.ref)}</code>` : ''}</li>`).join('')}</ul></details>`;
  }
  function controlsHTML(instance) {
    const data=instance.data, filters=data.filters;
    const groups=data.groups || [];
    const selectedSheet=instance.state.sheets.find(sheet=>sheet.name===filters.sheet);
    const allRows=(data.allRecords || []).filter(row=>filters.groupId==='all' || row.groupId===filters.groupId);
    const sortOptions=[['source:asc','ترتيب المصدر'],['name:asc','الاسم: أ–ي'],['duration:desc','المدة: من الأعلى'],['duration:asc','المدة: من الأقل'],...(data.sheet==='Sheet2' ? [['count:desc','الظهور: من الأعلى'],['seconds:desc','زمن العرض: من الأعلى'],['dailyHours:desc','الساعات اليومية: من الأعلى'],['monthlyHours:desc','الساعات الشهرية: من الأعلى']] : [])];
    return `<section class="md-filters" aria-label="فلاتر لوحة المتابعة"><div class="md-filter-top"><span class="md-filter-title">نطاق العرض</span><button type="button" class="md-reset" data-dash-reset>مسح جميع الفلاتر</button></div><div class="md-filter-grid"><label><span>ورقة المصدر</span><select data-dash-filter="sheet">${instance.state.sheets.map(sheet=>`<option value="${escape(sheet.name)}" ${filters.sheet===sheet.name ? 'selected' : ''}>${text(sheet.name)} · ${sheet.name==='Sheet2' ? 'حسابات العرض' : 'مدد المحتوى'}</option>`).join('')}</select></label><label class="md-filter-group"><span>المجموعة / الشاشة الواردة في المصدر</span><select data-dash-filter="groupId"><option value="all" ${filters.groupId==='all' ? 'selected' : ''}>جميع مجموعات الورقة</option>${groups.map(group=>`<option value="${escape(group.id)}" ${filters.groupId===group.id ? 'selected' : ''}>${text(group.title)}</option>`).join('')}</select></label><label class="md-filter-record"><span>المحتوى</span><select data-dash-filter="recordId"><option value="all">جميع محتويات النطاق</option>${allRows.map(record=>`<option value="${escape(record.id)}" ${filters.recordId===record.id ? 'selected' : ''}>${text(nameOf(record))} · ${text((record.inputRefs?.name || 'B'+record.sourceRow))}</option>`).join('')}</select></label><label class="md-search"><span>بحث</span><input type="search" data-dash-filter="search" placeholder="ابحث بالاسم أو مرجع الخلية…" value="${escape(filters.search)}" autocomplete="off"></label><label><span>ترتيب المحتوى</span><select data-dash-filter="sort">${sortOptions.map(([value,label])=>`<option value="${value}" ${filters.sort+':'+filters.direction===value ? 'selected' : ''}>${text(label)}</option>`).join('')}</select></label><label class="md-missing-filter"><input type="checkbox" data-dash-filter="missing" ${filters.missing ? 'checked' : ''}><span>السجلات ذات البيانات الناقصة فقط</span></label></div><p class="md-filter-note">${text(selectedSheet?.name || '')} · ${data.filters.groupId==='all' ? `${groups.length} مجموعات · ${data.sheet==='Sheet1' && data.coverage?.duration?.aggregatable===false ? 'مجموعات مستقلة؛ المدد لا تُجمع' : 'الأوقات تراكمية عبر المجموعات'}` : text(data.group?.sourceRange || '')} · فلاتر واحدة للبطاقات والرسوم والجدول</p></section>`;
  }
  function resultsHTML(instance) {
    const data=instance.data;
    const selected=data.records.find(row=>row.id===data.filters.recordId);
    const allowed=data.coverage?.duration?.aggregatable!==false && percentAllowed(data.records,'duration');
    const percentage=instance.mode==='percent' && allowed;
    const modeToggle=`<div class="md-mode-toggle" role="group" aria-label="طريقة عرض رسوم المحتوى"><button type="button" data-dash-mode="value" aria-pressed="${instance.mode==='value'}">القيم</button><button type="button" data-dash-mode="percent" aria-pressed="${instance.mode==='percent'}" ${!allowed ? 'disabled title="يتطلب مجموعًا موجبًا من قيم غير سالبة"' : ''}>النسب %</button></div>`;
    return `<div class="md-scope-line"><div><span class="md-scope-label">${data.filters.groupId==='all' ? (data.sheet==='Sheet1' && data.coverage?.duration?.aggregatable===false ? 'مجموعات مستقلة؛ المدد لا تُجمع' : 'تراكمي عبر المجموعات') : text(data.group?.title || 'النطاق المختار')}</span><span class="md-result-count"><b class="md-digits">${data.records.length}</b> سجل معروض</span></div>${selected ? `<button type="button" class="md-selection-chip" data-dash-record="${escape(selected.id)}">${text(nameOf(selected))}<span aria-hidden="true">×</span><span class="sr-only">إلغاء تحديد المحتوى</span></button>` : ''}</div>${summaryCards(data)}<p class="md-metric-note">${data.sheet==='Sheet2' ? 'C مدد المقاطع، F مرات الظهور، E زمن العرض، G وI ساعات العرض؛ لكل مؤشر وحدته ومعادلته الأصلية.' : 'العمود A محفوظ بوحدته الأصلية؛ لا يُحوّل إلى ثوانٍ أو دقائق لأن الوحدة غير معنْونة في المصدر.'} ${data.filters.groupId==='all' ? (data.sheet==='Sheet1' && data.coverage?.duration?.aggregatable===false ? 'تُعرض مدد كل مجموعة مستقلة؛ لا تُجمع وحدات غير مؤكدة.' : 'المجاميع عبر المجموعات تراكمية؛ لا تمثّل مدة تشغيل شاشة واحدة.') : ''}</p>${warningHTML(data)}<div class="md-content-grid"><section class="md-panel md-duration-panel">${metricHeader('مدد المحتوى','التوزيع داخل النطاق',percentage ? 'نسبة كل مدة من مجموع المدد المتاحة بعد التصفية.' : 'اسم كامل، مرجع واضح، ومدّة مستقلة لكل سجل.',modeToggle)}${horizontalChart(data.records,'duration',data,instance,{percentage})}<p class="md-chart-hint">اضغط أي محتوى لربط جميع الرسوم والجدول به. مرّر المؤشر أو استخدم Tab لعرض القيمة كاملة.</p></section>${shareChart(data)}${appearanceChart(data,instance)}${comparisonChart(data,instance)}</div>${linkedTable(data,instance)}<p class="md-bottom-note">مرجع اللوحة: ${text(instance.state.document?.sourceName || 'ملف Excel')} · إصدار <span class="md-digits">${Number(instance.state.document?.revision || 0)}</span>. يعرض الملخص القيم المتاحة دون تعديل بيانات المصدر أو استكمال الخلايا الناقصة.</p>`;
  }
  function getData(instance) {
    if (!global.MaestroDashboardData?.build) throw new Error('تعذر تحميل حسابات لوحة المتابعة. أعد تحميل الصفحة.');
    const data=global.MaestroDashboardData.build(instance.state,instance.filters);
    instance.filters={...data.filters};
    instance.data=data;
    if (data.sheet!=='Sheet2') instance.comparisonMetric='duration';
    if (instance.mode==='percent' && (data.coverage?.duration?.aggregatable===false || !percentAllowed(data.records,'duration'))) instance.mode='value';
    return data;
  }
  function rememberFocus(instance) {
    const active=global.document.activeElement;
    if (!active || !instance.container.contains(active)) return null;
    if (active.dataset.dashFilter) return {filter:active.dataset.dashFilter,start:active.selectionStart,end:active.selectionEnd};
    if (active.hasAttribute('data-dash-reset')) return {reset:true};
    const control=['dashRecord','dashGroup','dashMode'].find(key=>active.dataset[key]);
    return control ? {control,value:active.dataset[control],panel:active.closest('.md-panel')?.className || ''} : null;
  }
  function restoreFocus(instance,focus) {
    if (!focus) return;
    const control=focus.reset ? instance.container.querySelector('[data-dash-reset]') : focus.filter
      ? Array.from(instance.container.querySelectorAll('[data-dash-filter]')).find(node=>node.dataset.dashFilter===focus.filter)
      : Array.from(instance.container.querySelectorAll('button')).find(node=>node.dataset[focus.control]===focus.value && (node.closest('.md-panel')?.className || '')===focus.panel && node.tabIndex!==-1);
    if (!control) return;
    if (control===global.document.activeElement) return;
    control.focus({preventScroll:true});
    if (control.type==='search' && typeof focus.start==='number') control.setSelectionRange(focus.start,focus.end);
  }
  function update(instance,controls=true) {
    clearTimeout(instance.searchTimer);
    hideTooltip(instance);
    const focus=rememberFocus(instance);
    try {
      getData(instance);
      if (!instance.container.querySelector('.md-dashboard')) {
        instance.container.innerHTML=`<section class="md-dashboard" aria-label="لوحة متابعة المحتوى"><header class="md-page-heading"><div><span class="md-eyebrow">MAESTRO / CONTENT OVERVIEW</span><h1>لوحة متابعة المحتوى</h1><p>قراءة مترابطة لمدد المحتوى ومرات ظهوره، من بيانات Excel ومعادلاته الأصلية.</p></div><span class="md-live-label"><i aria-hidden="true"></i>البيانات المحفوظة</span></header><div data-dash-controls></div><div data-dash-results></div><div class="md-tooltip" id="md-tooltip-${instance.id}" role="tooltip" hidden></div></section>`;
        controls=true;
      }
      if (controls) instance.container.querySelector('[data-dash-controls]').innerHTML=controlsHTML(instance);
      instance.container.querySelector('[data-dash-results]').innerHTML=resultsHTML(instance);
      restoreFocus(instance,focus);
    } catch(error) {
      instance.container.innerHTML=`<section class="md-dashboard"><h1>لوحة متابعة المحتوى</h1><div class="notice warning">${text(error.message)}</div></section>`;
    }
  }
  function hideTooltip(instance) {
    if (instance.tooltipTarget) instance.tooltipTarget.removeAttribute('aria-describedby');
    instance.tooltipTarget=null;
    const node=instance.container.querySelector('.md-tooltip');
    if (node) node.hidden=true;
  }
  function showTooltip(instance,target) {
    const node=instance.container.querySelector('.md-tooltip');
    if (!node || !target) return;
    let data;
    try { data=JSON.parse(target.dataset.dashTooltip); } catch (_) { return; }
    if (instance.tooltipTarget && instance.tooltipTarget!==target) instance.tooltipTarget.removeAttribute('aria-describedby');
    instance.tooltipTarget=target;
    target.setAttribute('aria-describedby',node.id);
    node.innerHTML=`<strong>${text(data.name)}</strong><div class="md-tooltip-value"><span class="md-digits">${precise(data.value)}</span> ${text(data.unit)}</div>${data.ref ? `<code>${text(data.ref)}</code>` : ''}${finite(data.share) ? `<div>النسبة: <span class="md-digits">${precise(data.share)}%</span></div>` : ''}${finite(data.denominator) ? `<div>المقام: <span class="md-digits">${precise(data.denominator)}</span> ${text(data.denominatorUnit || data.unit)}</div>` : ''}${data.note ? `<p>${text(data.note)}</p>` : ''}`;
    node.hidden=false;
    const rect=target.getBoundingClientRect(),box=node.getBoundingClientRect();
    const left=Math.max(12,Math.min(global.innerWidth-box.width-12,(rect.left+rect.right-box.width)/2));
    let top=rect.top-box.height-10;
    if (top<12) top=rect.bottom+10;
    top=Math.max(12,Math.min(global.innerHeight-box.height-12,top));
    node.style.left=left+'px';node.style.top=top+'px';
  }
  function bind(instance) {
    const container=instance.container;
    const listen=(type,handler)=>{container.addEventListener(type,handler);instance.listeners.push([type,handler]);};
    listen('click',event=>{
      const target=event.target.closest('button');if(!target || !container.contains(target) || target.disabled)return;
      if(target.hasAttribute('data-dash-reset')){instance.filters=defaultFilters(instance.state);instance.mode='value';instance.comparisonMetric='duration';update(instance);}
      else if(target.dataset.dashRecord){instance.filters.recordId=instance.filters.recordId===target.dataset.dashRecord ? 'all' : target.dataset.dashRecord;update(instance);}
      else if(target.dataset.dashGroup){instance.filters.groupId=target.dataset.dashGroup;instance.filters.recordId='all';update(instance);}
      else if(target.dataset.dashMode){instance.mode=target.dataset.dashMode;update(instance,false);}
      else if(target.dataset.dashEdit && instance.state.canEdit && typeof instance.options.onEdit==='function'){hideTooltip(instance);instance.options.onEdit({sheet:instance.data.sheet,groupId:target.dataset.dashEditGroup,recordId:target.dataset.dashEdit});}
    });
    listen('change',event=>{
      const key=event.target.dataset.dashFilter;if(!key || key==='search')return;
      if(key==='comparisonMetric')instance.comparisonMetric=event.target.value;
      else if(key==='sort'){const [sort,direction]=event.target.value.split(':');instance.filters.sort=sort;instance.filters.direction=direction;}
      else if(key==='sheet'){instance.filters.sheet=event.target.value;const sheet=instance.state.sheets.find(item=>item.name===event.target.value);instance.filters.groupId=sheet?.groups?.[0]?.id || 'all';instance.filters.recordId='all';instance.filters.sort='source';instance.filters.direction='asc';}
      else if(key==='groupId'){instance.filters.groupId=event.target.value;instance.filters.recordId='all';}
      else instance.filters[key]=key==='missing' ? event.target.checked : event.target.value;
      update(instance);
    });
    listen('input',event=>{
      if(event.target.dataset.dashFilter!=='search')return;
      instance.filters.search=event.target.value;clearTimeout(instance.searchTimer);
      instance.searchTimer=setTimeout(()=>update(instance,false),120);
    });
    listen('pointerover',event=>{const target=event.target.closest('[data-dash-tooltip]');if(target && container.contains(target))showTooltip(instance,target);});
    listen('pointerout',event=>{const target=event.target.closest('[data-dash-tooltip]');if(target && !target.contains(event.relatedTarget) && global.document.activeElement!==target)hideTooltip(instance);});
    listen('focusin',event=>{const target=event.target.closest('[data-dash-tooltip]');if(target)showTooltip(instance,target);});
    listen('focusout',event=>{if(!instance.tooltipTarget?.contains(event.relatedTarget))hideTooltip(instance);});
    listen('keydown',event=>{if(event.key==='Escape')hideTooltip(instance);});
  }
  function render(container,state,options={}) {
    if(!container || !state)return;
    if(!state.sheets?.length){clear(container);container.innerHTML='<section class="md-dashboard"><div class="md-page-heading"><div><span class="md-eyebrow">MAESTRO / CONTENT OVERVIEW</span><h1>لوحة متابعة المحتوى</h1><p>استورد ملف Excel المرجعي لعرض الملخص والرسوم المرتبطة ببياناته.</p></div></div><div class="md-chart-empty">لا يوجد ملف محفوظ بعد.</div></section>';return;}
    let instance=instances.get(container);
    if(!instance){instance={id:++serial,container,state,options,filters:defaultFilters(state),mode:'value',comparisonMetric:'duration',listeners:[],tooltipTarget:null};instances.set(container,instance);bind(instance);}
    instance.state=state;instance.options=options;update(instance);
  }
  function clear(container) {
    for(const [node,instance] of instances){
      if(container && container!==node)continue;
      clearTimeout(instance.searchTimer);hideTooltip(instance);
      for(const [type,handler] of instance.listeners)node.removeEventListener(type,handler);
      if(node.querySelector('.md-dashboard'))node.replaceChildren();
      instance.state=null;instance.data=null;instance.filters=null;instances.delete(node);
    }
  }
  function refresh(state) {for(const instance of instances.values()){if(state)instance.state=state;if(instance.container.querySelector('.md-dashboard'))update(instance);}}
  global.MaestroDashboard=Object.freeze({render,refresh,clear});
})(window);
