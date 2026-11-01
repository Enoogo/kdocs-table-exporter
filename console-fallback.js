// ============================================================
// KDocs 表格导出器 - 控制台免安装版
// 用法：打开 kdocs/WPS 365 或 腾讯文档在线表格页面 -> F12 -> Console
//       粘贴本文件全部内容 -> 回车
// 结果：把「第一个工作表」导出为 TSV 文件下载（可改 CONFIG 参数）
// ============================================================
(function () {
  var CONFIG = { sheetIndex: 0, maxRows: 20000, maxCols: 256 };
  var isTencent = /(^|\.)docs\.qq\.com$/i.test(location.hostname);

  /* ---------- 金山 / WPS ---------- */
  function wpsSheets() {
    return window.APP && window.APP.workbook && window.APP.workbook._worksheets && window.APP.workbook._worksheets._sheets;
  }

  function wpsCellText(v) {
    if (v == null) return '';
    var p = v && v.private ? v.private : v;
    if (typeof p === 'string') return p;
    if (typeof p !== 'object') return String(p);
    if (p.str != null) return String(p.str);
    if (p.num != null) return String(p.num);
    if (p.value != null && typeof p.value !== 'object') return String(p.value);
    if (p.bool != null) return String(p.bool);
    if (p.err != null) return String(p.err);
    return '';
  }

  /* ---------- 腾讯文档 ---------- */
  function tencentSheets() {
    try {
      var wm = window.SpreadsheetApp && window.SpreadsheetApp.workbook && window.SpreadsheetApp.workbook.worksheetManager;
      if (!wm) return null;
      var list = null;
      try { list = wm.getSheetList ? wm.getSheetList() : null; } catch (e) {}
      if (!Array.isArray(list) || !list.length) list = wm.sheetList;
      return Array.isArray(list) && list.length ? list : null;
    } catch (e) { return null; }
  }

  function tencentCellText(ws, r, c) {
    try {
      var cell = ws.getCellDataAtPosition(r, c);
      if (!cell) return '';
      var fv = cell.formattedValue;
      if (fv && fv.value != null) {
        var s = String(fv.value);
        if (s !== '') return s;
      }
      var v = cell.value;
      if (v == null) return '';
      if (typeof v === 'object') {
        if (v.text != null) return String(v.text);
        if (v.value != null && typeof v.value !== 'object') return String(v.value);
        return '';
      }
      return String(v);
    } catch (e) { return ''; }
  }

  /* ---------- 统一入口 ---------- */
  var all = isTencent ? tencentSheets() : wpsSheets();
  if (!all || !all.length) {
    console.error('[导出失败] 未找到表格数据模型，请确认打开的是「表格」文档且已加载完成，然后刷新页面重试。');
    return;
  }

  var ws = all[CONFIG.sheetIndex] || all[0];
  var name = '';
  try {
    name = isTencent ? ws.getSheetName() : ws.sheetData.getName();
  } catch (e) {
    name = 'Sheet' + (CONFIG.sheetIndex + 1);
  }

  var lastRow = 0, lastCol = 0;
  if (isTencent) {
    try {
      var ur = ws.cellDataGrid && ws.cellDataGrid.usedRange;
      if (ur) { lastRow = ur.endRowIndex; lastCol = ur.endColIndex; }
      else { lastRow = ws.getRowCount() - 1; lastCol = ws.getColCount() - 1; }
    } catch (e) {}
  } else {
    try {
      var gc = ws.sheetData.private.gridCells;
      lastRow = gc.getLastRow();
      lastCol = gc.calcLastCol();
    } catch (e) {}
  }
  lastRow = Math.min(CONFIG.maxRows - 1, Math.max(0, lastRow | 0));
  lastCol = Math.min(CONFIG.maxCols - 1, Math.max(0, lastCol | 0));

  var grid = [], maxR = -1, maxC = -1;
  for (var r = 0; r <= lastRow; r++) {
    var row = [];
    for (var c = 0; c <= lastCol; c++) {
      var t = isTencent ? tencentCellText(ws, r, c) : '';
      if (!isTencent) {
        try { t = wpsCellText(ws.getCell(r, c).getValue()); } catch (e) {}
      }
      row.push(t);
      if (t) { if (r > maxR) maxR = r; if (c > maxC) maxC = c; }
    }
    grid.push(row);
  }

  var tsv = grid
    .slice(0, maxR + 1)
    .map(function (row) {
      return row.slice(0, maxC + 1).map(function (s) {
        return String(s).replace(/[\t\r\n]+/g, ' ').trim();
      }).join('\t');
    })
    .join('\n');

  var filename = document.title.replace(/[\\/:*?"<>|\r\n\t]+/g, '_') + '_' + name + '.tsv';
  var blob = new Blob(['\ufeff' + tsv], { type: 'text/plain;charset=utf-8' });
  var a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();

  console.log('[导出完成] ' + filename + '，共 ' + (maxR + 1) + ' 行 × ' + (maxC + 1) + ' 列');
  console.log('提示：改 CONFIG.sheetIndex 可导出其它工作表（0 = 第一个）。');
})();
