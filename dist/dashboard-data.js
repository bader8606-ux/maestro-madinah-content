/* Pure dashboard projections. Contains no workbook data, storage or networking.
 * Browser: MaestroDashboardData.build(state, filters).
 * Node ESM: import this file, then use globalThis.MaestroDashboardData.
 * CommonJS outside an ESM package: require(...) returns the same API.
 */
(function (root, factory) {
  'use strict';
  var api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  if (root) root.MaestroDashboardData = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  'use strict';

  var KEYS = ['duration', 'count', 'seconds', 'dailyHours', 'monthlyHours'];
  var METRIC_COLUMNS = {count: 'F', seconds: 'E', dailyHours: 'G', monthlyHours: 'I'};
  var MISSING = 'غير متوفر';
  function numeric(value) { return typeof value === 'number' && Number.isFinite(value); }
  function copy(object) { return Object.assign({}, object || {}); }
  function text(value) { return value === null || value === undefined ? '' : String(value); }
  function presentName(value) { return typeof value === 'string' && !!value.trim(); }

  function metadata(sheet) {
    var calculated = sheet === 'Sheet2';
    return {
      duration: {label: 'مجموع مدد المحتوى', unit: calculated ? 'ثانية' : 'وحدة غير محددة في المصدر', column: calculated ? 'C' : 'A', applicable: !!sheet, basis: calculated ? 'مدة المحتوى في العمود C' : 'القيم الأصلية في العمود A؛ لا تحويل لوحدتها'},
      count: {label: 'عدد مرات الظهور يوميًا', unit: calculated ? 'مرة/يوم' : MISSING, column: calculated ? 'F' : null, applicable: calculated, basis: calculated ? 'القيم الأصلية في F؛ الوحدة مستدلّة من E ÷ C' : MISSING},
      seconds: {label: 'ثواني الظهور اليومية', unit: calculated ? 'ثانية/يوم' : MISSING, column: calculated ? 'E' : null, applicable: calculated, basis: calculated ? 'القيم الأصلية في العمود E' : MISSING},
      dailyHours: {label: 'ساعات الظهور اليومية', unit: calculated ? 'ساعة/يوم' : MISSING, column: calculated ? 'G' : null, applicable: calculated, basis: calculated ? 'القيم الأصلية في G؛ لا تُستكمل الخلايا الخالية من E' : MISSING},
      monthlyHours: {label: 'ساعات الظهور الشهرية', unit: calculated ? 'ساعة/شهر' : MISSING, column: calculated ? 'I' : null, applicable: calculated, basis: calculated ? 'القيم الأصلية في I وفق أيام H لكل سجل؛ لا افتراض لعدد الأيام' : MISSING}
    };
  }

  function sourceRef(record, key, sheet) {
    if (key === 'duration') return record.inputRefs.duration || (sheet === 'Sheet2' ? 'C' : 'A') + record.sourceRow;
    if (sheet !== 'Sheet2') return null;
    return (METRIC_COLUMNS[key] || '') + record.sourceRow;
  }

  function normalizedRecord(record, group, sheet, order) {
    var result = Object.assign({}, record, {
      groupId: group.id, groupTitle: group.title, sourceRange: group.sourceRange,
      sheet: sheet, sourceOrder: order,
      inputRefs: copy(record.inputRefs), metrics: copy(record.metrics),
      formulas: copy(record.formulas), errors: copy(record.errors),
      missingOutputs: (record.missingOutputs || []).slice(),
      name: record.name === undefined ? null : record.name,
      duration: record.duration === undefined ? null : record.duration,
      days: record.days === undefined ? null : record.days,
      extra: record.extra === undefined ? null : record.extra
    });
    ['ratio', 'count', 'seconds', 'dailyHours', 'monthlyHours'].forEach(function (key) {
      result[key] = sheet === 'Sheet2' && !result.errors[key] && record.metrics && record.metrics[key] !== undefined ? record.metrics[key] : null;
      if (result.errors[key]) result.metrics[key] = null;
    });
    ['name', 'duration', 'days', 'extra'].forEach(function (key) { if (result.errors[key]) result[key] = null; });
    result.hasMissing = !presentName(result.name) || !numeric(result.duration) || result.duration <= 0 ||
      (sheet === 'Sheet2' && (!numeric(result.days) || ['ratio', 'count', 'seconds', 'dailyHours', 'monthlyHours'].some(function (key) { return !numeric(result[key]); }))) ||
      Object.keys(result.errors).length > 0;
    result.displayName = presentName(result.name) ? result.name : MISSING;
    result.sharePercent = null;
    result.shareDenominator = null;
    return result;
  }

  function aggregate(records, sheet, meta) {
    var totals = {}, coverage = {}, provenance = {};
    KEYS.forEach(function (key) {
      var known = [], missing = [], errors = [], sum = 0;
      // Table sorting must not change the order of floating-point summation.
      records.slice().sort(function (a, b) { return a.sourceOrder - b.sourceOrder; }).forEach(function (record) {
        var reference = sourceRef(record, key, sheet);
        if (meta[key].applicable && numeric(record[key])) { known.push(reference); sum += record[key]; }
        else { missing.push(reference); if (reference && record.errors[key]) errors.push(reference); }
      });
      totals[key] = meta[key].applicable && known.length ? sum : null;
      coverage[key] = {
        available: known.length, total: records.length, missing: missing.length, errors: errors.length,
        partial: known.length > 0 && known.length < records.length,
        complete: records.length > 0 && known.length === records.length,
        applicable: meta[key].applicable, aggregatable: meta[key].applicable
      };
      provenance[key] = {sheet: sheet, column: meta[key].column, unit: meta[key].unit,
        refs: known.filter(Boolean), missingRefs: missing.filter(Boolean), errorRefs: errors, aggregation: 'sum_available_only',
        basis: meta[key].basis, applicable: meta[key].applicable};
    });
    return {totals: totals, coverage: coverage, provenance: provenance};
  }

  function warningInGroup(warning, group) {
    if (warning.groupId) return warning.groupId === group.id;
    var range = /[A-Z]+(\d+):[A-Z]+(\d+)/i.exec(group.sourceRange || '');
    var reference = /^[A-Z]+(\d+)$/i.exec(warning.ref || '');
    return !!(range && reference && Number(reference[1]) >= Number(range[1]) && Number(reference[1]) <= Number(range[2]));
  }

  function searchable(record) {
    return [record.name, record.groupTitle, record.sourceRow, record.inputRefs.name,
      record.inputRefs.duration, record.inputRefs.days, record.duration, record.days,
      record.extra, record.ratio, record.count, record.seconds, record.dailyHours,
      record.monthlyHours].map(text).join(' ').toLocaleLowerCase('ar');
  }

  function build(data, input) {
    input = input || {};
    var sheets = data && Array.isArray(data.sheets) ? data.sheets : [];
    var defaultSheet = sheets.find(function (sheet) { return sheet.name === 'Sheet2'; }) || sheets[0] || null;
    var selectedSheet = sheets.find(function (sheet) { return sheet.name === input.sheet; }) || defaultSheet;
    var sheet = selectedSheet ? selectedSheet.name : null;
    var originalGroups = selectedSheet && Array.isArray(selectedSheet.groups) ? selectedSheet.groups : [];
    var groups = originalGroups.map(function (group) {
      return {id: group.id, title: group.title, hasTitle: !!group.hasTitle,
        sourceRange: group.sourceRange, recordCount: (group.records || []).length};
    });
    var requestedGroup = groups.find(function (group) { return group.id === input.groupId; });
    var groupId = input.groupId === 'all' ? 'all' : requestedGroup ? requestedGroup.id : groups[0] ? groups[0].id : null;
    var selectedGroups = originalGroups.filter(function (group) { return groupId === 'all' || group.id === groupId; });
    var order = 0;
    var allRecords = originalGroups.flatMap(function (group) {
      return (group.records || []).map(function (record) { return normalizedRecord(record, group, sheet, order++); });
    });
    var groupIds = selectedGroups.map(function (group) { return group.id; });
    var scopedRecords = allRecords.filter(function (record) { return groupIds.includes(record.groupId); });
    var recordId = input.recordId && input.recordId !== 'all' && scopedRecords.some(function (record) { return record.id === input.recordId; }) ? input.recordId : 'all';
    var search = typeof input.search === 'string' ? input.search.trim() : '';
    var query = search.toLocaleLowerCase('ar');
    var sort = ['source', 'name', 'duration', 'count', 'seconds', 'dailyHours', 'monthlyHours', 'days'].includes(input.sort) ? input.sort : 'source';
    var direction = input.direction === 'desc' ? 'desc' : 'asc';
    var filters = {sheet: sheet, groupId: groupId, recordId: recordId, search: search, sort: sort, direction: direction, missing: input.missing === true};
    var records = scopedRecords.filter(function (record) {
      return (recordId === 'all' || record.id === recordId) && (!query || searchable(record).includes(query)) && (!filters.missing || record.hasMissing);
    });
    var collator = new Intl.Collator('ar', {numeric: true, sensitivity: 'base'});
    records.sort(function (a, b) {
      var comparison = 0;
      if (sort === 'source') comparison = a.sourceOrder - b.sourceOrder;
      else if (sort === 'name') comparison = collator.compare(text(a.name), text(b.name));
      else {
        var aMissing = !numeric(a[sort]), bMissing = !numeric(b[sort]);
        if (aMissing !== bMissing) return aMissing ? 1 : -1;
        if (!aMissing) comparison = a[sort] - b[sort];
      }
      return (direction === 'desc' ? -comparison : comparison) || a.sourceOrder - b.sourceOrder;
    });
    var meta = metadata(sheet), summary = aggregate(records, sheet, meta);
    var cumulative = selectedGroups.length > 1;
    var scope = {
      sheet: sheet, groupIds: groupIds, selectedGroupCount: selectedGroups.length,
      sourceRecordCount: allRecords.length, scopedRecordCount: scopedRecords.length,
      selectedRecordCount: records.length,
      aggregateBasis: cumulative ? 'cumulative_across_groups' : 'single_group',
      label: cumulative ? 'تراكمي عبر المجموعات' : (selectedGroups[0] ? selectedGroups[0].title : MISSING),
      elapsedTime: false, screenMultiplierApplied: false
    };
    var notices = [];
    if (cumulative) notices.push('الإجماليات تراكمية عبر المجموعات المحددة، وليست مدة فعلية منقضية أو مضروبة في عدد شاشات.');
    if (sheet === 'Sheet1') {
      notices.push('وحدة مدد Sheet1 وأساسها غير محددين في المصدر؛ تُعرض القيم الأصلية دون تحويل إلى ثوانٍ أو ساعات.');
      if (cumulative) {
        var reason = 'لا يُجمع العمود A عبر مجموعات Sheet1 لأن وحداتها وأسسها غير مؤكدة في المصدر.';
        summary.totals.duration = null;
        summary.coverage.duration.aggregatable = false;
        summary.provenance.duration.aggregation = 'not_aggregated_across_unconfirmed_groups';
        summary.provenance.duration.reason = reason;
        notices.push(reason);
      }
    }
    var denominator = sheet === 'Sheet2' ? summary.totals.seconds : null;
    var shares = records.map(function (record) {
      var percent = sheet === 'Sheet2' && numeric(record.seconds) && numeric(denominator) && denominator > 0 ? record.seconds / denominator * 100 : null;
      record.sharePercent = percent;
      record.shareDenominator = denominator;
      return {id: record.id, recordId: record.id, name: record.name, groupId: record.groupId,
        sourceRow: record.sourceRow, value: sheet === 'Sheet2' ? record.seconds : null,
        percent: percent, denominator: denominator, ref: sheet === 'Sheet2' ? 'E' + record.sourceRow : null,
        basis: 'currently_filtered_available_E', unit: sheet === 'Sheet2' ? 'ثانية/يوم' : MISSING};
    });
    var sheetWarnings = (data && Array.isArray(data.warnings) ? data.warnings : []).filter(function (warning) { return warning.sheet === sheet; });
    var warnings = sheetWarnings.filter(function (warning) { return selectedGroups.some(function (group) { return warningInGroup(warning, group); }); }).map(copy);
    var comparisons = selectedGroups.map(function (group) {
      var groupRecords = records.filter(function (record) { return record.groupId === group.id; });
      var groupSummary = aggregate(groupRecords, sheet, meta);
      return {id: group.id, title: group.title, hasTitle: !!group.hasTitle, sourceRange: group.sourceRange,
        recordCount: (group.records || []).length, selectedRecordCount: groupRecords.length,
        originalTotals: copy(group.totals), originalTotalRefs: copy(group.totalRefs),
        selectedTotals: groupSummary.totals, coverage: groupSummary.coverage,
        provenance: groupSummary.provenance,
        warnings: sheetWarnings.filter(function (warning) { return warningInGroup(warning, group); }).map(copy)};
    });
    return {
      sheet: sheet, group: groupId === 'all' ? null : groups.find(function (group) { return group.id === groupId; }) || null,
      groups: groups, records: records, selectedRows: records, allRecords: allRecords, flatRecords: allRecords,
      totals: summary.totals, coverage: summary.coverage, metricMeta: meta,
      recordCount: records.length, totalRecords: allRecords.length, scope: scope,
      comparisons: comparisons, comparisonFilters: copy(filters), shares: shares,
      warnings: warnings, notices: notices, filters: filters, missingLabel: MISSING,
      provenance: {
        sheet: sheet, groupIds: groupIds, recordIds: records.map(function (record) { return record.id; }),
        sourceRanges: selectedGroups.map(function (group) { return group.sourceRange; }),
        metrics: summary.provenance, aggregation: 'available_source_values_only',
        originalTotals: comparisons.map(function (group) { return {id: group.id, values: group.originalTotals, refs: group.originalTotalRefs}; }),
        missingValuesFilled: false, crossSheetAggregation: false, originalWarningsPreserved: true,
        shareBasis: 'currently_filtered_available_E'
      }
    };
  }

  return Object.freeze({build: build, missingLabel: MISSING});
});
