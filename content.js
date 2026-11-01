/**
 * KDocs 表格导出器 - content script（隔离世界）
 * 职责：注入页面世界脚本（金山/WPS）、转发弹窗命令、执行文件下载。
 * 腾讯文档：inject.js 由 manifest 以 world:MAIN 直接注入，见 manifest.json。
 */
(function () {
  var isTencent = /(^|\.)docs\.qq\.com$/i.test(location.hostname);

  /* ---- 注入页面世界脚本（金山文档 / WPS 365） ---- */
  function injectScript(file, marker) {
    if (document.querySelector('script[data-' + marker + ']')) return;
    var s = document.createElement('script');
    s.src = chrome.runtime.getURL(file);
    s.dataset[marker] = '1';
    s.onload = function () { s.remove(); };
    (document.head || document.documentElement).appendChild(s);
  }

  function inject() {
    if (!isTencent) {
      // 金山/WPS：沿用脚本标签注入（probe 必须最先注入，以便在页面脚本建立 WebSocket 之前挂钩）
      injectScript('probe.js', 'kdteProbe');
      injectScript('inject.js', 'kdte');
      return;
    }
    // 腾讯文档：manifest 已用 world:MAIN 注入 inject.js。
    // 若浏览器不支持 world:MAIN（Chrome < 111），再尝试脚本标签注入作为兜底。
    setTimeout(function () {
      var ok = false;
      try { ok = document.documentElement.getAttribute('data-kdte-main-world') === '1'; } catch (e) {}
      if (!ok) injectScript('inject.js', 'kdte');
    }, 1200);
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
    // JSON 不加 BOM，避免部分解析器 JSON.parse 失败；CSV/TSV 加 BOM 方便 Excel 识别 UTF-8
    var prefix = /json/i.test(mime || '') ? '' : '\ufeff';
    var blob = new Blob([prefix + content], { type: (mime || 'text/plain') + ';charset=utf-8' });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = filename || 'export.txt';
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
          var ext = msg.format === 'csv' ? 'csv' : msg.format === 'json' ? 'json' : 'tsv';
          res.files.forEach(function (f, i) {
            var name = (msg.fileNames && msg.fileNames[i]) || ('export_' + (i + 1) + '.' + ext);
            setTimeout(function () { download(name, f.content, mime); }, i * 400);
          });
          sendResponse({ ok: true, files: res.files.map(function (f, i) {
            var name = (msg.fileNames && msg.fileNames[i]) || ('export_' + (i + 1) + '.' + ext);
            return { name: name, rowCount: f.rowCount, colCount: f.colCount, bytes: f.content.length };
          }) });
        })
        .catch(function (e) { sendResponse({ ok: false, error: String(e.message || e) }); });
      return true;
    }
  });
})();
