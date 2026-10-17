/**
 * KDocs 表格导出器 - 页面世界注入脚本
 *
 * 原理：WPS/金山文档在线表格用 canvas 渲染，DOM 里没有文字节点，
 * 但所有单元格数据都保存在页面的 JS 数据模型里：
 *   APP.workbook._worksheets._sheets[i].sheetData.private.gridCells
 *   -> worksheet.getCell(row, col).getValue().private.str / .num
 * 本脚本在"页面世界"里直接读取该模型，绕过任何复制限制。
 */
(function () {
  if (window.__KDTE_INJECTED__) return;
  window.__KDTE_INJECTED__ = true;

  var MAX_ROWS_DEFAULT = 20000;
  var MAX_COLS_DEFAULT = 256;

  /* ---------- 定位数据模型（带多个兼容入口） ---------- */
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

  /* ---------- 抽取单表为二维数组（原始文本，保留换行） ---------- */
  function extractSheet(ws, idx, capRows, capCols) {
    var b = usedBounds(ws, capRows, capCols);
    var grid = [];
    var maxR = -1, maxC = -1;
    for (var r = 0; r <= b.rows; r++) {
      var row = [];
      for (var c = 0; c <= b.cols; c++) {
        var t = getCellText(ws, r, c);
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
      name: sheetName(ws, idx),
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
      if (d.cmd === 'listSheets') {
        var sheets = getSheetObjects();
        if (!sheets) {
          reply(reqId, {
            ok: false,
            error: '未找到表格数据模型。请确认：1) 打开的是「表格」文档（不是文字/演示）；2) 表格已加载完成（稍等几秒）；3) 刷新页面后重试。'
          });
          return;
        }
        var list = [];
        for (var i = 0; i < sheets.length; i++) {
          var name = sheetName(sheets[i], i);
          if (/^WpsReserved/i.test(name)) continue; // 跳过引擎保留表
          var b = usedBounds(sheets[i], MAX_ROWS_DEFAULT, MAX_COLS_DEFAULT);
          list.push({ index: i, name: name, approxRows: b.rows + 1, approxCols: b.cols + 1 });
        }
        reply(reqId, { ok: true, sheets: list, title: document.title });
        return;
      }

      if (d.cmd === 'export' || d.cmd === 'export-download') {
        var all = getSheetObjects();
        if (!all) {
          reply(reqId, { ok: false, error: '未找到表格数据模型（同上）' });
          return;
        }
        var capRows = Math.max(1, Math.min(200000, parseInt(d.maxRows, 10) || MAX_ROWS_DEFAULT));
        var capCols = Math.max(1, Math.min(1024, parseInt(d.maxCols, 10) || MAX_COLS_DEFAULT));
        var idxs = Array.isArray(d.sheetIndexes) && d.sheetIndexes.length
          ? d.sheetIndexes
          : all.map(function (_, i) { return i; }).filter(function (i) { return !/^WpsReserved/i.test(sheetName(all[i], i)); });

        var extracted = idxs.map(function (i) {
          return extractSheet(all[i], i, capRows, capCols);
        });

        if (d.cmd === 'export') {
          reply(reqId, {
            ok: true,
            format: d.format || 'tsv',
            sheets: extracted.map(function (s) {
              return { name: s.name, rowCount: s.rowCount, colCount: s.colCount };
            }),
            text: serialize(extracted, d.format || 'tsv')
          });
          return;
        }
        // export-download：返回序列化结果 + 每个 sheet 一份内容，供 content script 落盘
        reply(reqId, {
          ok: true,
          format: d.format || 'tsv',
          files: extracted.map(function (s) {
            var body = d.format === 'json' ? toJSON(extracted) : serialize([s], d.format || 'tsv');
            return { name: s.name, rowCount: s.rowCount, colCount: s.colCount, content: body };
          })
        });
        return;
      }

      if (d.cmd === 'probe-status') {
        reply(reqId, { ok: true, installed: !!window.__KDTE_PROBE__, stats: window.__KDTE_PROBE__ || null });
        return;
      }

      if (d.cmd === 'probe-reset') {
        var pr = window.__KDTE_PROBE__;
        if (!pr) {
          reply(reqId, { ok: false, error: '探针未安装：请确认插件是 v1.2.0、已在 chrome://extensions 重新加载，并硬刷新页面（Ctrl+Shift+R）。' });
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
        var trg = { ok: true, results: [] };
        try {
          var tsheets = getSheetObjects() || [];
          for (var ti = 0; ti < tsheets.length; ti++) {
            try {
              var lp = tsheets[ti].loadSheetData && tsheets[ti].loadSheetData();
              trg.results.push({ sheet: ti, ok: true, isPromise: !!(lp && lp.then) });
            } catch (e2) {
              trg.results.push({ sheet: ti, ok: false, error: String(e2) });
            }
          }
        } catch (e3) {
          trg.ok = false;
          trg.error = String(e3);
        }
        reply(reqId, trg);
        return;
      }

      reply(reqId, { ok: false, error: '未知命令: ' + d.cmd });
    } catch (err) {
      reply(reqId, { ok: false, error: String((err && err.message) || err) });
    }
  });

  window.postMessage({ source: 'kdte-page', type: 'ready' }, '*');
})();
