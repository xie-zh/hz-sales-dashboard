(function () {
  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>]/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c];
    });
  }
  function show(s) {
    if (!s || !s.stale) return;
    var b = document.createElement('div');
    b.id = 'stale-banner';
    b.style.cssText = 'position:fixed;left:0;right:0;top:0;z-index:99999;' +
      'background:#b05a52;color:#fff;padding:7px 14px;text-align:center;' +
      'font:13px/1.7 -apple-system,"Microsoft YaHei",sans-serif;' +
      'box-shadow:0 2px 10px rgba(0,0,0,.28)';
    b.innerHTML = '⚠ 数据未更新：最近一次同步失败' +
      (s.last_error ? '（' + esc(s.last_error) + '）' : '') +
      ' · 当前显示 ' + esc(s.last_success || '上一轮') + ' 的数据' +
      (s.updated_at ? ' · 检查于 ' + esc(s.updated_at) : '');
    document.body.appendChild(b);
    document.body.style.paddingTop = (b.offsetHeight + 4) + 'px';
  }
  function load() {
    fetch('status.json?cb=' + Date.now(), { cache: 'no-store' })
      .then(function (r) { return r.ok ? r.json() : null; })
      .then(show).catch(function () {});
  }
  if (document.readyState !== 'loading') load();
  else document.addEventListener('DOMContentLoaded', load);
})();
