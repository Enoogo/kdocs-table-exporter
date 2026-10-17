// ============================================================
// KDocs 表格导出器 - 控制台免安装版
// 用法：打开 kdocs/WPS 365 在线表格页面 -> 按 F12 -> Console
//       粘贴本文件全部内容 -> 回车
// 结果：把「第一个工作表」导出为 TSV 文件下载（可改参数）
// ============================================================
(function () {
  var CONFIG = { sheetIndex: 0, maxRows: 20000, maxCols: 256 };

  function sheets() {
    return window.APP && APP.workbook && APP.workbook._worksheets && APP.workbook._worksheets._sheets;
  }

  function cellText(v) {
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

  var all = sheets();
  if (!all || !all.length) {
    console.error('[导出失败] 未找到表格数据模型，请确认打开的是「表格」文档且已加载完成，然后刷新页面重试。');
    return;
  }

  var ws = all[CONFIG.sheetIndex] || all[0];
  var name = '';
  try { name = ws.sheetData.getName(); } catch (e) { name = 'Sheet' + (CONFIG.sheetIndex + 1); }

  var lastRow = 0, lastCol = 0;
  try {
    var gc = ws.sheetData.private.gridCells;
    lastRow = gc.getLastRow();
    lastCol = gc.calcLastCol();
  } catch (e) {}
  lastRow = Math.min(CONFIG.maxRows - 1, lastRow | 0);
  lastCol = Math.min(CONFIG.maxCols - 1, lastCol | 0);

  var grid = [], maxR = -1, maxC = -1;
  for (var r = 0; r <= lastRow; r++) {
    var row = [];
    for (var c = 0; c <= lastCol; c++) {
      var t = '';
      try { t = cellText(ws.getCell(r, c).getValue()); } catch (e) {}
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
