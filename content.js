/**
 * KDocs 表格导出器 - content script（隔离世界）
 * 职责：注入 inject.js（页面世界）、转发弹窗命令、执行文件下载。
 */
(function () {
  /* ---- 注入页面世界脚本 ---- */
  function injectScript(file, marker) {
    if (document.querySelector('script[data-' + marker + ']')) return;
    var s = document.createElement('script');
    s.src = chrome.runtime.getURL(file);
    s.dataset[marker] = '1';
    s.onload = function () { s.remove(); };
    (document.head || document.documentElement).appendChild(s);
  }
  function inject() {
    // 探针必须最先注入（document_start），以便在页面脚本建立 WebSocket 之前挂钩
    injectScript('probe.js', 'kdteProbe');
    injectScript('inject.js', 'kdte');
  }
  inject();

  /* ---- 请求桥 ---- */
  var pending = {};
  var seq = 0;

  window.addEventListener('message', function (e) {
    if (e.source !== window) return;
    var d = e.data;
    if (!d || d.source !== 'kdte-page' || !d.reqId) return;
    var cb = pending[d.reqId];
    if (cb) {
      delete pending[d.reqId];
      cb(d);
    }
  });

  function askPage(payload, timeoutMs) {
    return new Promise(function (resolve, reject) {
      var reqId = 'kdte' + (++seq) + '_' + Date.now();
      pending[reqId] = resolve;
      var msg = { source: 'kdte-cmd', reqId: reqId };
      for (var k in payload) msg[k] = payload[k];
      window.postMessage(msg, '*');
      setTimeout(function () {
        if (pending[reqId]) {
          delete pending[reqId];
          reject(new Error('表格响应超时。请确认表格已加载完成，刷新页面后重试。'));
        }
      }, timeoutMs || 60000);
    });
  }

  /* ---- 下载（Blob + a[download]，无需 downloads 权限） ---- */
  function download(filename, content, mime) {
    var blob = new Blob(['\ufeff' + content], { type: (mime || 'text/plain') + ';charset=utf-8' });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.style.display = 'none';
    document.body.appendChild(a);
    a.click();
    setTimeout(function () {
      a.remove();
      URL.revokeObjectURL(url);
    }, 30000);
  }

  /* ---- 弹窗消息 ---- */
  chrome.runtime.onMessage.addListener(function (msg, sender, sendResponse) {
    if (!msg || msg.source !== 'kdte-ext') return;

    if (msg.cmd === 'ping') {
      sendResponse({ ok: true, title: document.title, href: location.href });
      return;
    }

    if (msg.cmd === 'probe-status' || msg.cmd === 'probe-reset' || msg.cmd === 'probe-trigger') {
      askPage({ cmd: msg.cmd }, 30000).then(sendResponse).catch(function (e) {
        sendResponse({ ok: false, error: String(e.message || e) });
      });
      return true;
    }

    if (msg.cmd === 'listSheets') {
      askPage({ cmd: 'listSheets' }).then(sendResponse).catch(function (e) {
        sendResponse({ ok: false, error: String(e.message || e) });
      });
      return true;
    }

    if (msg.cmd === 'export') {
      askPage({ cmd: 'export', format: msg.format, sheetIndexes: msg.sheetIndexes, maxRows: msg.maxRows, maxCols: msg.maxCols }, 120000)
        .then(sendResponse)
        .catch(function (e) { sendResponse({ ok: false, error: String(e.message || e) }); });
      return true;
    }

    if (msg.cmd === 'export-download') {
      askPage({ cmd: 'export-download', format: msg.format, sheetIndexes: msg.sheetIndexes, maxRows: msg.maxRows, maxCols: msg.maxCols }, 120000)
        .then(function (res) {
          if (!res.ok) { sendResponse(res); return; }
          var mime = msg.format === 'csv' ? 'text/csv' : msg.format === 'json' ? 'application/json' : 'text/plain';
          res.files.forEach(function (f, i) {
            setTimeout(function () { download(msg.fileNames[i], f.content, mime); }, i * 400);
          });
          sendResponse({ ok: true, files: res.files.map(function (f, i) {
            return { name: msg.fileNames[i], rowCount: f.rowCount, colCount: f.colCount, bytes: f.content.length };
          }) });
        })
        .catch(function (e) { sendResponse({ ok: false, error: String(e.message || e) }); });
      return true;
    }
  });
})();
