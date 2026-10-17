/**
 * KDocs 表格导出器 - 网络探针（诊断模式）
 *
 * 目的：判断敏感字段的脱敏（●●●●●●●）发生在哪一层：
 *   - 若服务器下发的数据里就有圆点、且搜不到身份证/手机号形态字符串  → 服务端脱敏（真值从未下发）
 *   - 若下发数据里存在身份证/手机号形态字符串                        → 客户端安装时脱码（真值到过浏览器）
 *
 * 严格只做统计：只累计「出现次数」和打码预览（前3位+****+后2位），
 * 不保存、不输出任何完整号码或单元格原文。
 * 结果挂在 window.__KDTE_PROBE__ 上，供诊断读取。
 */
(function () {
  if (window.__KDTE_PROBE__) return;

  var stats = {
    probeVersion: 1,
    startedAt: Date.now(),
    startedBeforeScripts: document.readyState === 'loading',
    wsFrames: 0,
    wsBinaryFrames: 0,
    wsGzipFrames: 0,
    wsBytes: 0,
    fetchCount: 0,
    fetchBytes: 0,
    xhrCount: 0,
    xhrBytes: 0,
    dotsInPayloads: 0,
    idLikeHits: 0,
    phoneLikeHits: 0,
    idPreviews: [],
    phonePreviews: [],
    markerPayloads: 0,
    payloadSamples: []
  };
  window.__KDTE_PROBE__ = stats;

  var ID_RE = /[1-9]\d{5}(?:19|20)\d{2}(?:0[1-9]|1[0-2])(?:0[1-9]|[12]\d|3[01])\d{3}[\dXx]/g;
  var PHONE_RE = /1[3-9]\d{9}/g;

  function scan(text, src) {
    if (typeof text !== 'string' || !text) return;
    if (src === 'ws') stats.wsBytes += text.length;
    else if (src === 'fetch') stats.fetchBytes += text.length;
    else if (src === 'xhr') stats.xhrBytes += text.length;

    var dots = text.split('●●●●●●●').length - 1;
    if (dots > 0) stats.dotsInPayloads += dots;

    var ids = text.match(ID_RE);
    if (ids && ids.length) {
      stats.idLikeHits += ids.length;
      for (var i = 0; i < Math.min(3, ids.length) && stats.idPreviews.length < 20; i++) {
        stats.idPreviews.push(ids[i].slice(0, 3) + '****' + ids[i].slice(-2));
      }
    }
    var ph = text.match(PHONE_RE);
    if (ph && ph.length) {
      stats.phoneLikeHits += ph.length;
      for (var j = 0; j < Math.min(3, ph.length) && stats.phonePreviews.length < 20; j++) {
        stats.phonePreviews.push(ph[j].slice(0, 3) + '****' + ph[j].slice(-2));
      }
    }
    if (text.indexOf('千佛山校区') >= 0 || text.indexOf('模板表') >= 0) {
      stats.markerPayloads++;
      if (stats.payloadSamples.length < 10) {
        stats.payloadSamples.push({ src: src, len: text.length, hasDots: dots > 0, hasId: !!(ids && ids.length) });
      }
    }
  }

  function decodeBytes(bytes, src) {
    try {
      if (bytes.length > 1 && bytes[0] === 0x1f && bytes[1] === 0x8b && typeof DecompressionStream === 'function') {
        if (src === 'ws') stats.wsGzipFrames++;
        var stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream('gzip'));
        new Response(stream).arrayBuffer().then(function (ab) {
          scan(new TextDecoder('utf-8', { fatal: false }).decode(new Uint8Array(ab)), src);
        }).catch(function () {});
        return;
      }
    } catch (e) {}
    scan(new TextDecoder('utf-8', { fatal: false }).decode(bytes), src);
  }

  function handlePayload(data, src) {
    try {
      if (typeof data === 'string') { scan(data, src); return; }
      if (typeof ArrayBuffer !== 'undefined' && data instanceof ArrayBuffer) { decodeBytes(new Uint8Array(data), src); return; }
      if (typeof Uint8Array !== 'undefined' && data instanceof Uint8Array) { decodeBytes(data, src); return; }
      if (data && typeof data.arrayBuffer === 'function') {
        data.arrayBuffer().then(function (ab) { decodeBytes(new Uint8Array(ab), src); }).catch(function () {});
      }
    } catch (e) {}
  }

  /* ---- WebSocket：拦截所有新建连接的下行帧 ---- */
  try {
    var OrigWS = window.WebSocket;
    if (OrigWS) {
      var Wrapped = function (url, protocols) {
        var ws = (protocols !== undefined) ? new OrigWS(url, protocols) : new OrigWS(url);
        try {
          ws.addEventListener('message', function (ev) {
            stats.wsFrames++;
            if (typeof ev.data !== 'string') stats.wsBinaryFrames++;
            handlePayload(ev.data, 'ws');
          });
        } catch (e) {}
        return ws;
      };
      Wrapped.prototype = OrigWS.prototype;
      Wrapped.CONNECTING = OrigWS.CONNECTING;
      Wrapped.OPEN = OrigWS.OPEN;
      Wrapped.CLOSING = OrigWS.CLOSING;
      Wrapped.CLOSED = OrigWS.CLOSED;
      window.WebSocket = Wrapped;
    }
  } catch (e) {}

  /* ---- fetch / XHR：只扫描数据类响应，跳过静态资源 ---- */
  function interesting(u) {
    if (!u) return false;
    u = String(u);
    if (/\.(js|css|png|jpg|jpeg|gif|svg|woff2?|ttf|ico|webp)(\?|$)/i.test(u)) return false;
    return true;
  }

  try {
    var origFetch = window.fetch;
    if (origFetch) {
      window.fetch = function () {
        try {
          var u = (arguments[0] && arguments[0].url) ? arguments[0].url : arguments[0];
          if (!interesting(u)) return origFetch.apply(this, arguments);
          stats.fetchCount++;
          var p = origFetch.apply(this, arguments);
          p.then(function (res) {
            try { res.clone().text().then(function (t) { scan(t, 'fetch'); }).catch(function () {}); } catch (e) {}
          }).catch(function () {});
          return p;
        } catch (e) {
          return origFetch.apply(this, arguments);
        }
      };
    }
  } catch (e) {}

  try {
    var origOpen = XMLHttpRequest.prototype.open;
    var origSend = XMLHttpRequest.prototype.send;
    XMLHttpRequest.prototype.open = function (method, url) {
      this.__kdte_url = url;
      return origOpen.apply(this, arguments);
    };
    XMLHttpRequest.prototype.send = function () {
      try {
        if (interesting(this.__kdte_url)) {
          var self = this;
          stats.xhrCount++;
          self.addEventListener('load', function () {
            var t = '';
            try { t = self.responseText; } catch (e) { return; }
            scan(t, 'xhr');
          });
        }
      } catch (e) {}
      return origSend.apply(this, arguments);
    };
  } catch (e) {}
})();
