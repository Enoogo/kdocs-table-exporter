/** KDocs 表格导出器 - 弹窗逻辑 */
var tabId = null;
var sheets = [];
var docTitle = 'kdocs';
var platform = 'kdocs';

function log(msg, isErr) {
  var el = document.getElementById('log');
  el.textContent = msg;
  el.className = isErr ? 'err' : '';
}

function send(msg) {
  return new Promise(function (resolve) {
    chrome.tabs.sendMessage(tabId, msg, function (res) {
      if (chrome.runtime.lastError) {
        resolve({ ok: false, error: '无法连接页面：' + chrome.runtime.lastError.message + '。请刷新该在线表格页面后重试。' });
        return;
      }
      resolve(res || { ok: false, error: '页面无响应' });
    });
  });
}

function sanitize(name) {
  return String(name || 'sheet').replace(/[\\/:*?"<>|\r\n\t]+/g, '_').slice(0, 80) || 'sheet';
}

function escHtml(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
  });
}

function selectedIndexes() {
  return Array.prototype.slice
    .call(document.querySelectorAll('#sheetList input[type=checkbox]:checked'))
    .map(function (i) { return parseInt(i.value, 10); });
}

function renderSheets() {
  var box = document.getElementById('sheetList');
  if (!sheets.length) {
    box.innerHTML = '<span class="muted">未发现工作表</span>';
    return;
  }
  box.innerHTML = sheets
    .map(function (s) {
      return (
        '<div class="sheet-item">' +
        '<input type="checkbox" id="s' + s.index + '" value="' + s.index + '" checked>' +
        '<label for="s' + s.index + '" title="' + escHtml(s.name) + '">' + escHtml(s.name) + '</label>' +
        '<span class="muted">≈' + s.approxRows + '×' + s.approxCols + '</span>' +
        '</div>'
      );
    })
    .join('');
}

function platformLabel() {
  return platform === 'tencent' ? '腾讯文档' : '金山文档/WPS 365';
}

async function loadSheets() {
  log('正在扫描工作表…');
  var res = await send({ source: 'kdte-ext', cmd: 'listSheets' });
  if (!res.ok) {
    log(res.error, true);
    return;
  }
  sheets = res.sheets || [];
  platform = res.platform || 'kdocs';
  docTitle = sanitize(res.title);
  document.getElementById('docInfo').textContent = docTitle + ' · ' + platformLabel() + ' · 共 ' + sheets.length + ' 个工作表';
  renderSheets();
  log(sheets.length ? '扫描完成，请选择工作表后导出。' : '未发现可导出的工作表。', !sheets.length);
}

function buildFileName(sheetName, format) {
  var ext = format === 'csv' ? 'csv' : format === 'json' ? 'json' : 'tsv';
  return sanitize(docTitle) + '_' + sanitize(sheetName) + '.' + ext;
}

async function doDownload() {
  var idxs = selectedIndexes();
  var format = document.getElementById('format').value;
  if (!idxs.length) { log('请先勾选至少一个工作表', true); return; }

  var maxRows = document.getElementById('maxRows').value;
  var maxCols = document.getElementById('maxCols').value;

  var names = idxs.map(function (i) {
    var s = sheets.filter(function (x) { return x.index === i; })[0] || { name: 'Sheet' + i };
    return buildFileName(s.name, format);
  });
  // 多表 JSON 会合并成一个文件
  var fileNames = (format === 'json' && idxs.length > 1) ? [sanitize(docTitle) + '.json'] : names;

  log('正在提取数据…（大表格可能需要几秒）');
  document.getElementById('btnDownload').disabled = true;
  var res = await send({
    source: 'kdte-ext',
    cmd: 'export-download',
    format: format,
    sheetIndexes: idxs,
    fileNames: fileNames,
    maxRows: maxRows,
    maxCols: maxCols
  });
  document.getElementById('btnDownload').disabled = false;

  if (!res.ok) { log(res.error, true); return; }
  var files = res.files || [];
  log(
    '✅ 已下载 ' + files.length + ' 个文件：\n' +
    files.map(function (f) {
      return '· ' + f.name + '（' + f.rowCount + ' 行 × ' + f.colCount + ' 列，' + Math.round(f.bytes / 1024) + ' KB）';
    }).join('\n')
  );
}

async function doCopy() {
  var idxs = selectedIndexes();
  if (!idxs.length) { log('请先勾选至少一个工作表', true); return; }
  log('正在提取数据…');
  var res = await send({
    source: 'kdte-ext',
    cmd: 'export',
    format: 'tsv',
    sheetIndexes: [idxs[0]],
    maxRows: document.getElementById('maxRows').value,
    maxCols: document.getElementById('maxCols').value
  });
  if (!res.ok) { log(res.error, true); return; }
  try {
    await navigator.clipboard.writeText(res.text);
    log('✅ 已复制 ' + res.sheets[0].name + '（' + res.sheets[0].rowCount + ' 行 × ' + res.sheets[0].colCount + ' 列）到剪贴板，可直接粘贴到 Excel。');
  } catch (e) {
    log('复制失败：' + e, true);
  }
}

/* ---------- 金山/WPS：网络探针诊断 ---------- */
function probeConcl(st) {
  if (!st) return '探针未安装。请确认插件已更新、在 chrome://extensions 点过「重新加载」、页面已硬刷新（Ctrl+Shift+R）。';
  if (st.markerPayloads === 0) {
    return '⚠️ 还没抓到表格数据载荷（可能是离线缓存恢复的数据）。\n请点「触发数据加载」，或在表格里滚动到底部、切换一次工作表，再点「读取探针统计」。';
  }
  var real = st.idLikeHits + st.phoneLikeHits;
  if (real > 0 && st.dotsInPayloads > 0) {
    return '🔎 结论：数据包里同时有真实号码形态和 ●●●●●●● → 客户端打码（真值到过浏览器，可以进一步提取）。';
  }
  if (real > 0) {
    return '🔎 结论：数据包里有真实号码形态、未见圆点 → 客户端打码，真值在下发数据里。';
  }
  if (st.dotsInPayloads > 0) {
    return '🔎 结论：数据包里只有 ●●●●●●●、搜不到任何号码形态 → 服务端脱敏，真值从未下发到浏览器，客户端无解。';
  }
  return '⚠️ 载荷里既没圆点也没号码形态，数据可能以编码/二进制格式下发，需要进一步抓包分析。';
}

function renderNetworkProbe(st) {
  if (!st) {
    return '探针未安装。请确认插件已更新、在 chrome://extensions 点过「重新加载」、页面已硬刷新（Ctrl+Shift+R）。';
  }
  return [
    '运行时长：' + Math.round((Date.now() - st.startedAt) / 1000) + ' 秒',
    'WebSocket 帧：' + st.wsFrames + '（二进制 ' + st.wsBinaryFrames + '，gzip ' + st.wsGzipFrames + '，字符数 ' + st.wsBytes + '）',
    'fetch：' + st.fetchCount + ' 次 / ' + st.fetchBytes + ' 字符',
    'XHR：' + st.xhrCount + ' 次 / ' + st.xhrBytes + ' 字符',
    '表格数据载荷：' + st.markerPayloads + ' 个',
    '载荷中 ●●●●●●● 出现：' + st.dotsInPayloads + ' 次',
    '身份证形态字符串：' + st.idLikeHits + ' 个' + (st.idPreviews && st.idPreviews.length ? '（预览 ' + st.idPreviews.join('、') + '）' : ''),
    '手机号形态字符串：' + st.phoneLikeHits + ' 个' + (st.phonePreviews && st.phonePreviews.length ? '（预览 ' + st.phonePreviews.join('、') + '）' : ''),
    '',
    probeConcl(st)
  ].join('\n');
}

/* ---------- 腾讯文档：表格模型诊断 ---------- */
function modelConcl(st) {
  if (!st) return '未读取到模型统计。';
  if (st.maskedWithRawValue > 0) {
    return '🔎 结论：显示值被打码，但模型 value 里能读到非打码原值 → 属于显示层/客户端打码，真值到过浏览器。';
  }
  if (st.maskedCells > 0) {
    return '🔎 结论：模型里的 value 本身就是打码的 → 打码发生在数据下发/建模阶段，客户端模型里没有读到非打码原值。';
  }
  return '✅ 未发现打码单元格。';
}

function renderModelProbe(st) {
  if (!st) return '未读取到模型统计。';
  var lines = [
    '诊断模式：表格模型（腾讯文档）',
    '工作表：' + st.sheets + ' 个',
    '扫描单元格：' + st.cellsScanned + '，非空：' + st.nonEmptyCells,
    '显示值含打码（****/●●●）：' + st.maskedCells + ' 个',
    '底层 value 与显示值不同：' + st.valueDiffersFromDisplay + ' 个',
    '打码但底层有非打码原值：' + st.maskedWithRawValue + ' 个'
  ];
  if (st.truncated) {
    lines.push('⚠️ 表格较大，诊断只扫描了前 5000 行 × 64 列。');
  }
  if (st.samples && st.samples.length) {
    lines.push('', '示例（仅打码预览）：');
    st.samples.forEach(function (s) {
      lines.push('· ' + s.sheet + ' R' + s.row + 'C' + s.col + '：' + s.display + ' / ' + s.rawPreview);
    });
  }
  lines.push('', modelConcl(st));
  return lines.join('\n');
}

function renderProbe(res) {
  if (res.mode === 'model') return renderModelProbe(res.stats);
  if (!res.installed) {
    return '探针未安装。请确认插件已更新、在 chrome://extensions 点过「重新加载」、页面已硬刷新（Ctrl+Shift+R）。';
  }
  return renderNetworkProbe(res.stats);
}

async function probeStatus() {
  var out = document.getElementById('probeOut');
  out.textContent = '读取中…';
  var res = await send({ source: 'kdte-ext', cmd: 'probe-status' });
  if (!res.ok) { out.textContent = res.error; return; }
  out.textContent = renderProbe(res);
}

async function probeTrigger() {
  var out = document.getElementById('probeOut');
  var res = await send({ source: 'kdte-ext', cmd: 'probe-trigger' });
  if (!res.ok) { out.textContent = res.error; return; }
  if (res.skipped) {
    out.textContent = '✅ 当前平台的数据已在内存模型中，无需触发加载。请直接点「读取探针统计」。';
    return;
  }
  out.textContent = '✅ 已触发 ' + res.results.length + ' 个工作表重新加载数据。\n等 2~3 秒后点「读取探针统计」。';
}

async function probeReset() {
  var out = document.getElementById('probeOut');
  var res = await send({ source: 'kdte-ext', cmd: 'probe-reset' });
  if (!res.ok) { out.textContent = res.error; return; }
  out.textContent = res.note === 'model' ? '✅ 模型诊断无需清零。' : '✅ 已清零统计。';
}

document.addEventListener('DOMContentLoaded', function () {
  var b1 = document.getElementById('btnProbeStatus');
  if (b1) b1.addEventListener('click', probeStatus);
  var b2 = document.getElementById('btnProbeTrigger');
  if (b2) b2.addEventListener('click', probeTrigger);
  var b3 = document.getElementById('btnProbeReset');
  if (b3) b3.addEventListener('click', probeReset);
});

document.addEventListener('DOMContentLoaded', async function () {
  var tabs = await chrome.tabs.query({ active: true, currentWindow: true });
  if (!tabs.length) { log('未找到活动标签页', true); return; }
  tabId = tabs[0].id;

  document.getElementById('btnDownload').addEventListener('click', doDownload);
  document.getElementById('btnCopy').addEventListener('click', doCopy);
  document.getElementById('btnRefresh').addEventListener('click', loadSheets);
  document.getElementById('selAll').addEventListener('click', function (e) {
    e.preventDefault();
    document.querySelectorAll('#sheetList input').forEach(function (i) { i.checked = true; });
  });
  document.getElementById('selNone').addEventListener('click', function (e) {
    e.preventDefault();
    document.querySelectorAll('#sheetList input').forEach(function (i) { i.checked = false; });
  });

  var url = tabs[0].url || '';
  if (!/kdocs\.cn|wps\.cn|docs\.qq\.com/.test(url)) {
    document.getElementById('docInfo').textContent = '当前页面不是金山文档/WPS 365/腾讯文档';
    log('请先打开一个金山文档/WPS 365 或 腾讯文档的「表格」文档。', true);
    return;
  }
  await loadSheets();
});
