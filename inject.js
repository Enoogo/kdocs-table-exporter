/**
 * KDocs 表格导出器 - 页面世界注入脚本
 *
 * 支持两个平台：
 *   1) 金山文档 / WPS 365（*.kdocs.cn、*.wps.cn）
 *      APP.workbook._worksheets._sheets[i].sheetData.private.gridCells
 *      -> worksheet.getCell(row, col).getValue().private.str / .num
 *   2) 腾讯文档在线表格（docs.qq.com/sheet/...）
 *      SpreadsheetApp.workbook.worksheetManager.getSheetList()
 *      -> sheet.getCellDataAtPosition(row, col).formattedValue.value / .value
 *
 * 本脚本在"页面世界"里直接读取该模型，绕过复制限制，全程只读。
 */
(function () {
  if (window.__KDTE_INJECTED__) return;
  window.__KDTE_INJECTED__ = true;

  var MAX_ROWS_DEFAULT = 20000;
  var MAX_COLS_DEFAULT = 256;

  /* ---------- 平台识别 ---------- */
  function isTencentDocs() {
    return /(^|\.)docs\.qq\.com$/i.test(location.hostname);
  }

  function getDocTitle() {
    var t = '';
    try { t = document.title || ''; } catch (e) {}
    if (t) return t;
    try {
      if (isTencentDocs() && window.docInfo && window.docInfo.padInfo && window.docInfo.padInfo.padTitle) {
        return String(window.docInfo.padInfo.padTitle);
      }
    } catch (e) {}
    return isTencentDocs() ? '腾讯文档' : 'kdocs';
  }

  /* ============================================================
   * 金山文档 / WPS 365 数据模型
   * ============================================================ */
  function getSheetObjects() {
    try {
      var sheets = window.APP && window.APP.workbook && window.APP.workbook._worksheets && window.APP.workbook._worksheets._sheets;
      if (Array.isArray(sheets) && sheets.length) return sheets;
    } catch (e) {}
    // 兜底：在 window 上搜索形状相同的 workbook 对象（应对引擎版本差异）
    try {
      for (var k in window) {
        try {
          var o = window[k];
          if (o && o.workbook && o.workbook._worksheets && Array.isArray(o.workbook._worksheets._sheets)) {
            return o.workbook._worksheets._sheets;
          }
        } catch (e) {}
      }
    } catch (e) {}
    return null;
  }

  function sheetName(ws, idx) {
    try {
      var n = ws.sheetData && ws.sheetData.getName && ws.sheetData.getName();
      if (n) return String(n);
    } catch (e) {}
    try {
      var n2 = ws.getName && ws.getName();
      if (n2) return String(n2);
    } catch (e) {}
    try {
      var n3 = ws._coreObject && ws._coreObject.name;
      if (n3 && !/^WorksheetObject/.test(n3)) return String(n3);
    } catch (e) {}
    return 'Sheet' + (idx + 1);
  }

  /* ---------- 单元格取值（str / num / bool / err / 其它私有字段） ---------- */
  function cellText(v) {
    if (v === null || v === undefined) return '';
    var p = v && v.private ? v.private : v;
    if (typeof p === 'string') return p;
    if (typeof p !== 'object') return String(p);
    if (p.str !== undefined && p.str !== null) return String(p.str);
    if (p.num !== undefined && p.num !== null) return String(p.num);
    if (p.value !== undefined && p.value !== null && typeof p.value !== 'object') return String(p.value);
    if (p.bool !== undefined && p.bool !== null) return String(p.bool);
    if (p.err !== undefined && p.err !== null) return String(p.err);
    for (var i = 0; i < Object.keys(p).length; i++) {
      var key = Object.keys(p)[i];
      if (key === 'tokenMajorType' || key === 'isSqLeading') continue;
      var x = p[key];
      if (typeof x === 'string' || typeof x === 'number' || typeof x === 'boolean') return String(x);
    }
    return '';
  }

  function getCellText(ws, r, c) {
    try {
      return cellText(ws.getCell(r, c).getValue());
    } catch (e) {
      return '';
    }
  }

  /* ---------- 有效范围 ---------- */
  function usedBounds(ws, capRows, capCols) {
    var lastRow = 0, lastCol = 0;
    try {
      var gc = ws.sheetData.private.gridCells;
      lastRow = gc.getLastRow();
      lastCol = gc.calcLastCol();
    } catch (e) {}
    lastRow = Math.min(capRows - 1, lastRow | 0);
    lastCol = Math.min(capCols - 1, lastCol | 0);
    return { rows: Math.max(lastRow, 0), cols: Math.max(lastCol, 0) };
  }

  /* ============================================================
   * 腾讯文档在线表格数据模型
   * ============================================================ */
  function getTencentSheetObjects() {
    try {
      var app = window.SpreadsheetApp;
      var wm = app && app.workbook && app.workbook.worksheetManager;
      if (wm) {
        var list = null;
        try { list = wm.getSheetList ? wm.getSheetList() : null; } catch (e) {}
        if (!Array.isArray(list) || !list.length) list = wm.sheetList;
        if (Array.isArray(list) && list.length) return list;
      }
    } catch (e) {}
    // 兜底：在 window 上搜索带 worksheetManager 的对象（应对引擎版本差异）
    try {
      for (var k in window) {
        try {
          var o = window[k];
          var wm2 = o && o.workbook && o.workbook.worksheetManager;
          if (wm2) {
            var l2 = null;
            try { l2 = wm2.getSheetList ? wm2.getSheetList() : null; } catch (e) {}
            if (!Array.isArray(l2) || !l2.length) l2 = wm2.sheetList;
            if (Array.isArray(l2) && l2.length) return l2;
          }
        } catch (e) {}
      }
    } catch (e) {}
    return null;
  }

  function tencentSheetName(ws, idx) {
    try {
      var n = ws.getSheetName && ws.getSheetName();
      if (n) return String(n);
    } catch (e) {}
    try {
      var n2 = ws._AnT;
      if (n2) return String(n2);
    } catch (e) {}
    return 'Sheet' + (idx + 1);
  }

  /* 腾讯文档单元格的底层值（不做格式化） */
  function tencentRawValue(cell) {
    if (!cell) return '';
    var v = cell.value;
    if (v === null || v === undefined) return '';
    if (typeof v === 'object') {
      if (v.text !== undefined && v.text !== null) return String(v.text);
      if (v.value !== undefined && v.value !== null && typeof v.value !== 'object') return String(v.value);
      return '';
    }
    return String(v);
  }

  /* 腾讯文档单元格的显示值（日期、百分比等按表格显示格式导出） */
  function tencentCellText(ws, r, c) {
    try {
      var cell = ws.getCellDataAtPosition(r, c);
      if (!cell) return '';
      var fv = cell.formattedValue;
      if (fv && fv.value !== undefined && fv.value !== null) return String(fv.value);
      return tencentRawValue(cell);
    } catch (e) {
      return '';
    }
  }

  function tencentUsedBounds(ws, capRows, capCols) {
    var lastRow = -1, lastCol = -1, hasRange = false;
    try {
      var ur = ws.cellDataGrid && ws.cellDataGrid.usedRange;
      if (ur) {
        hasRange = true;
        lastRow = ur.endRowIndex;
        lastCol = ur.endColIndex;
      }
    } catch (e) {}
    if (hasRange) {
      if (lastRow < 0 || lastCol < 0) return { rows: 0, cols: 0 };
    } else {
      try { lastRow = ws.getRowCount() - 1; } catch (e) { lastRow = 0; }
      try { lastCol = ws.getColCount() - 1; } catch (e) { lastCol = 0; }
    }
    lastRow = Math.min(capRows - 1, Math.max(0, lastRow | 0));
    lastCol = Math.min(capCols - 1, Math.max(0, lastCol | 0));
    return { rows: lastRow, cols: lastCol };
  }

  /* ============================================================
   * 平台适配层
   * ============================================================ */
  function getPlatform() {
    if (isTencentDocs()) {
      return {
        id: 'tencent',
        label: '腾讯文档',
        getSheets: getTencentSheetObjects,
        sheetName: tencentSheetName,
        cellText: tencentCellText,
        bounds: tencentUsedBounds,
        skipSheet: function () { return false; },
        notFound: '未找到腾讯文档表格数据模型。请确认：1) 打开的是「在线表格」文档（不是文字/幻灯片/智能表格）；2) 表格已加载完成（稍等几秒）；3) 刷新页面后重试。'
      };
    }
    return {
      id: 'kdocs',
      label: '金山文档/WPS 365',
      getSheets: getSheetObjects,
      sheetName: sheetName,
      cellText: getCellText,
      bounds: usedBounds,
      skipSheet: function (name) { return /^WpsReserved/i.test(name); },
      notFound: '未找到表格数据模型。请确认：1) 打开的是「表格」文档（不是文字/演示）；2) 表格已加载完成（稍等几秒）；3) 刷新页面后重试。'
    };
  }

  /* ---------- 抽取单表为二维数组（原始文本，保留换行） ---------- */
  function extractSheet(p, ws, idx, capRows, capCols) {
    var b = p.bounds(ws, capRows, capCols);
    var grid = [];
    var maxR = -1, maxC = -1;
    for (var r = 0; r <= b.rows; r++) {
      var row = [];
      for (var c = 0; c <= b.cols; c++) {
        var t = p.cellText(ws, r, c);
        row.push(t);
        if (t) {
          if (r > maxR) maxR = r;
          if (c > maxC) maxC = c;
        }
      }
      grid.push(row);
    }
    var rows = grid.slice(0, maxR + 1).map(function (row) {
      return row.slice(0, maxC + 1);
    });
    return {
      name: p.sheetName(ws, idx),
      rows: rows,
      rowCount: rows.length,
      colCount: maxC + 1
    };
  }

  /* ---------- 序列化 ---------- */
  function flat(s) {
    return String(s == null ? '' : s).replace(/[\t\r\n]+/g, ' ').trim();
  }

  function toTSV(sheet) {
    return sheet.rows
      .map(function (row) {
        return row.map(flat).join('\t');
      })
      .join('\n');
  }

  function csvCell(s) {
    s = String(s == null ? '' : s);
    return /[",\r\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  }

  function toCSV(sheet) {
    return sheet.rows
      .map(function (row) {
        return row.map(csvCell).join(',');
      })
      .join('\r\n');
  }

  function toJSON(sheets) {
    var out = {};
    sheets.forEach(function (s) {
      out[s.name] = s.rows;
    });
    return JSON.stringify(out, null, 2);
  }

  function serialize(sheets, format) {
    if (format === 'json') return toJSON(sheets);
    if (format === 'csv') return sheets.map(toCSV).join('\r\n');
    return sheets.map(toTSV).join('\n');
  }

  /* ============================================================
   * 腾讯文档模型诊断：判断打码发生在显示层还是模型层
   * 只统计次数，最多保存 5 条打码预览（前3位****后2位），不保存完整原文。
   * ============================================================ */
  function modelProbe(p) {
    var stats = {
      mode: 'model',
      platform: p.id,
      sheets: 0,
      cellsScanned: 0,
      nonEmptyCells: 0,
      maskedCells: 0,
      maskedWithRawValue: 0,
      valueDiffersFromDisplay: 0,
      truncated: false,
      samples: []
    };
    var sheets = p.getSheets();
    if (!sheets) return { error: p.notFound };
    var MASK_RE = /●{3,}|\*{3,}/;
    for (var si = 0; si < sheets.length; si++) {
      var ws = sheets[si];
      stats.sheets++;
      var b = p.bounds(ws, MAX_ROWS_DEFAULT, MAX_COLS_DEFAULT);
      var maxR = Math.min(b.rows, 5000);
      var maxC = Math.min(b.cols, 64);
      if (maxR < b.rows || maxC < b.cols) stats.truncated = true;
      for (var r = 0; r <= maxR; r++) {
        for (var c = 0; c <= maxC; c++) {
          var cell = null;
          try { cell = ws.getCellDataAtPosition(r, c); } catch (e) { continue; }
          if (!cell) continue;
          stats.cellsScanned++;
          var fv = cell.formattedValue;
          var display = (fv && fv.value !== undefined && fv.value !== null) ? String(fv.value) : '';
          if (display === '') display = tencentRawValue(cell);
          var raw = tencentRawValue(cell);
          if (!display && !raw) continue;
          stats.nonEmptyCells++;
          var displayMasked = MASK_RE.test(display);
          var rawMasked = MASK_RE.test(raw);
          if (displayMasked) stats.maskedCells++;
          if (raw !== display) stats.valueDiffersFromDisplay++;
          if (displayMasked && !rawMasked && raw && raw !== display) {
            stats.maskedWithRawValue++;
            if (stats.samples.length < 5) {
              stats.samples.push({
                sheet: p.sheetName(ws, si),
                row: r + 1,
                col: c + 1,
                display: display.slice(0, 40),
                rawPreview: raw.length > 6 ? raw.slice(0, 3) + '****' + raw.slice(-2) : '***'
              });
            }
          }
        }
      }
    }
    return stats;
  }

  /* ---------- 消息协议（页面世界 <-> content script） ---------- */
  function reply(reqId, payload) {
    var msg = { source: 'kdte-page', reqId: reqId };
    for (var k in payload) msg[k] = payload[k];
    window.postMessage(msg, '*');
  }

  window.addEventListener('message', function (e) {
    if (e.source !== window) return;
    var d = e.data;
    if (!d || d.source !== 'kdte-cmd') return;

    var reqId = d.reqId;
    try {
      var platform = getPlatform();

      if (d.cmd === 'listSheets') {
        var sheets = platform.getSheets();
        if (!sheets) {
          reply(reqId, { ok: false, error: platform.notFound });
          return;
        }
        var list = [];
        for (var i = 0; i < sheets.length; i++) {
          var name = platform.sheetName(sheets[i], i);
          if (platform.skipSheet(name)) continue; // 跳过 WPS 引擎保留表
          var b = platform.bounds(sheets[i], MAX_ROWS_DEFAULT, MAX_COLS_DEFAULT);
          list.push({ index: i, name: name, approxRows: b.rows + 1, approxCols: b.cols + 1 });
        }
        reply(reqId, { ok: true, sheets: list, title: getDocTitle(), platform: platform.id });
        return;
      }

      if (d.cmd === 'export' || d.cmd === 'export-download') {
        var all = platform.getSheets();
        if (!all) {
          reply(reqId, { ok: false, error: platform.notFound });
          return;
        }
        var capRows = Math.max(1, Math.min(200000, parseInt(d.maxRows, 10) || MAX_ROWS_DEFAULT));
        var capCols = Math.max(1, Math.min(1024, parseInt(d.maxCols, 10) || MAX_COLS_DEFAULT));
        var idxs = Array.isArray(d.sheetIndexes) && d.sheetIndexes.length
          ? d.sheetIndexes
          : all.map(function (_, i) { return i; }).filter(function (i) { return !platform.skipSheet(platform.sheetName(all[i], i)); });

        var extracted = idxs.map(function (i) {
          return extractSheet(platform, all[i], i, capRows, capCols);
        });

        if (d.cmd === 'export') {
          reply(reqId, {
            ok: true,
            platform: platform.id,
            format: d.format || 'tsv',
            sheets: extracted.map(function (s) {
              return { name: s.name, rowCount: s.rowCount, colCount: s.colCount };
            }),
            text: serialize(extracted, d.format || 'tsv')
          });
          return;
        }

        // export-download：返回序列化结果 + 每个 sheet 一份内容，供 content script 落盘
        var fmt = d.format || 'tsv';
        if (fmt === 'json' && extracted.length > 1) {
          // 多个工作表合并为一个 JSON 文件
          reply(reqId, {
            ok: true,
            platform: platform.id,
            format: fmt,
            files: [{
              name: 'all',
              rowCount: extracted.reduce(function (a, s) { return a + s.rowCount; }, 0),
              colCount: extracted.reduce(function (a, s) { return Math.max(a, s.colCount); }, 0),
              content: toJSON(extracted)
            }]
          });
          return;
        }
        reply(reqId, {
          ok: true,
          platform: platform.id,
          format: fmt,
          files: extracted.map(function (s) {
            var body = fmt === 'json' ? toJSON([s]) : serialize([s], fmt);
            return { name: s.name, rowCount: s.rowCount, colCount: s.colCount, content: body };
          })
        });
        return;
      }

      if (d.cmd === 'probe-status') {
        if (platform.id === 'tencent') {
          var mp = modelProbe(platform);
          if (mp.error) {
            reply(reqId, { ok: false, error: mp.error });
            return;
          }
          reply(reqId, { ok: true, mode: 'model', platform: platform.id, stats: mp });
          return;
        }
        reply(reqId, { ok: true, mode: 'network', platform: platform.id, installed: !!window.__KDTE_PROBE__, stats: window.__KDTE_PROBE__ || null });
        return;
      }

      if (d.cmd === 'probe-reset') {
        if (platform.id === 'tencent') {
          reply(reqId, { ok: true, note: 'model' });
          return;
        }
        var pr = window.__KDTE_PROBE__;
        if (!pr) {
          reply(reqId, { ok: false, error: '探针未安装：请确认插件已更新、已在 chrome://extensions 重新加载，并硬刷新页面（Ctrl+Shift+R）。' });
          return;
        }
        ['wsFrames', 'wsBinaryFrames', 'wsGzipFrames', 'wsBytes', 'fetchCount', 'fetchBytes', 'xhrCount', 'xhrBytes', 'dotsInPayloads', 'idLikeHits', 'phoneLikeHits', 'markerPayloads'].forEach(function (k) { pr[k] = 0; });
        pr.idPreviews = [];
        pr.phonePreviews = [];
        pr.payloadSamples = [];
        reply(reqId, { ok: true });
        return;
      }

      if (d.cmd === 'probe-trigger') {
        var tsheets = platform.getSheets() || [];
        var trg = { ok: true, results: [], skipped: false };
        var anyLoaded = false;
        for (var ti = 0; ti < tsheets.length; ti++) {
          var ws = tsheets[ti];
          if (typeof ws.loadSheetData === 'function') {
            anyLoaded = true;
            try {
              var lp = ws.loadSheetData();
              trg.results.push({ sheet: ti, ok: true, isPromise: !!(lp && lp.then) });
            } catch (e2) {
              trg.results.push({ sheet: ti, ok: false, error: String(e2) });
            }
          } else {
            trg.results.push({ sheet: ti, ok: true, skipped: true });
          }
        }
        trg.skipped = !anyLoaded;
        reply(reqId, trg);
        return;
      }

      if (d.cmd === 'ping-page') {
        reply(reqId, { ok: true, platform: platform.id, hasModel: !!platform.getSheets() });
        return;
      }

      reply(reqId, { ok: false, error: '未知命令: ' + d.cmd });
    } catch (err) {
      reply(reqId, { ok: false, error: String((err && err.message) || err) });
    }
  });

  // 供 content script 判断页面世界脚本是否已注入（Chrome < 111 时用于兜底）
  try { document.documentElement.setAttribute('data-kdte-main-world', '1'); } catch (e) {}
  window.postMessage({ source: 'kdte-page', type: 'ready' }, '*');
})();
