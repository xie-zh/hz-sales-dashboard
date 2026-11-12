/* ============================================================================
   员工通道 v3：侧栏导航 + 顶部轮播 + 公告列表 + 六大分区卡片
   ----------------------------------------------------------------------------
   数据全部来自 site.config.js：
     staffBanners[] —— 顶部轮播（图片 / 标题 / 副标题）
     portal[]       —— 六大分区与全部入口（取自原看板导航页 nav.html）
     notices[]      —— 公告（简报 / 通知）
   // Motion Prompt #074 · MIT License (c) 2026 Shohei Ohtaki · 滚动淡入
   // Motion Prompt #081 · MIT License (c) 2026 Shohei Ohtaki · 悬停抬升
   // https://showheyohtaki.github.io/motion-prompt/
   ============================================================================ */
(function () {
  'use strict';
  var S = window.SITE || {};
  var cats = S.portal || [];
  var notices = (S.notices || []).slice().sort(function (a, b) {
    return (b.date + b.time).localeCompare(a.date + a.time);
  });

  var $ = function (s, r) { return (r || document).querySelector(s); };
  var $$ = function (s, r) { return Array.prototype.slice.call((r || document).querySelectorAll(s)); };
  var esc = function (v) {
    return String(v == null ? '' : v).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  };
  var pad = function (n) { return String(n).padStart(2, '0'); };
  var totalCount = cats.reduce(function (a, c) { return a + c.files.length; }, 0);
  var onlineCount = cats.reduce(function (a, c) {
    return a + c.files.filter(function (f) { return !!f.cover; }).length;
  }, 0);

  /* ---------------------------------------------------------------- 侧栏统计 */
  var stats = $('#staffStats');
  if (stats) {
    stats.innerHTML =
      '<div><b data-cu="' + totalCount + '">0</b><span>入口总数</span></div>' +
      '<div><b data-cu="' + onlineCount + '">0</b><span>在线看板</span></div>' +
      '<div><b data-cu="' + cats.length + '">0</b><span>业务分区</span></div>';
  }

  /* #007 Count-up：0 → 目标值，1 − (1 − p)^4 先快后稳；用时间戳而非 rAF，
     iframe / 后台标签页里 rAF 会被冻结，这里保证一定能落到终值 */
  function countUp(el, target, dur, delay) {
    var start = Date.now() + (delay || 0);
    var timer = setInterval(function () {
      var p = Math.min(1, (Date.now() - start) / dur);
      if (p <= 0) return;
      el.textContent = Math.round(target * (1 - Math.pow(1 - p, 4)));
      if (p >= 1) clearInterval(timer);
    }, 40);
    setTimeout(function () { el.textContent = target; }, (delay || 0) + dur + 140);
  }
  if (stats) {
    $$('[data-cu]', stats).forEach(function (el, i) {
      countUp(el, parseInt(el.getAttribute('data-cu'), 10) || 0, 860, 200 + i * 110);
    });
  }

  /* ---------------------------------------------------------------- 公告弹窗 */
  var modal = $('#noticeModal'), mbody = $('#noticeModalBody');
  function openModal(html) {
    if (!modal || !mbody) return;
    mbody.innerHTML = html;
    modal.classList.add('is-open');
    modal.setAttribute('aria-hidden', 'false');
    document.body.style.overflow = 'hidden';
  }
  function closeModal() {
    if (!modal) return;
    modal.classList.remove('is-open');
    modal.setAttribute('aria-hidden', 'true');
    document.body.style.overflow = '';
  }
  function openNotice(n) {
    if (!n) return;
    openModal(
      '<div class="nt-head"><h2 id="ntTitle">' + esc(n.title) + '</h2>' +
        '<div class="nt-meta">' + esc(n.date) + ' ' + esc(n.time) + ' · ' +
        (n.cat === 'brief' ? '数据简报' : '重要通知') + '</div></div>' +
      '<div class="nt-body">' +
        (n.content || []).map(function (p) { return '<p>' + esc(p) + '</p>'; }).join('') +
      '</div>');
  }
  function openList() {
    openModal(
      '<div class="nt-head"><h2 id="ntTitle">全部公告</h2>' +
        '<div class="nt-meta">共 ' + notices.length + ' 条 · 按时间倒序</div></div>' +
      '<div class="nt-list">' + notices.map(function (n, i) {
        return '<button type="button" class="nt-row" data-nl="' + i + '">' +
          '<time>' + esc(n.date.slice(5)) + ' ' + esc(n.time) + '</time>' +
          '<span class="nt-cat">' + (n.cat === 'brief' ? '简报' : '通知') + '</span>' +
          '<b>' + esc(n.title) + '</b></button>';
      }).join('') + '</div>');
  }
  if (modal) {
    document.addEventListener('click', function (e) {
      var t = e.target;
      if (!t || !t.closest) return;
      if (t.closest('[data-modal-close]')) { closeModal(); return; }
      var row = t.closest('[data-nl]');
      if (row) openNotice(notices[parseInt(row.getAttribute('data-nl'), 10)]);
    });
    document.addEventListener('keydown', function (e) {
      if (e.key === 'Escape' && modal.classList.contains('is-open')) closeModal();
    });
  }

  /* ---------------------------------------------------------------- 顶部轮播 */
  var banners = S.staffBanners || [];
  var bTrack = $('#bnTrack'), bDots = $('#bnDots');
  if (bTrack && banners.length) {
    bTrack.innerHTML = banners.map(function (b, i) {
      var media = b.image
        ? '<img src="' + esc(b.image) + '" alt=""' +
          (b.pos ? ' style="object-position:' + esc(b.pos) + '"' : '') + '>'
        : '<div style="width:100%;height:100%;background:linear-gradient(135deg,#eef3fb,#e2eafa)"></div>';
      return '<div class="bn-slide" data-bn="' + i + '">' + media +
        '<div class="bn-cap">' +
          '<b>' + esc(b.title) + '</b>' +
          (b.sub ? '<span>' + esc(b.sub) + '</span>' : '') +
        '</div></div>';
    }).join('');

    if (bDots) {
      bDots.innerHTML = banners.map(function (_, i) {
        return '<button type="button" data-bdot="' + i + '" aria-label="第 ' + (i + 1) + ' 张"></button>';
      }).join('');
    }

    var bi = 0, bTimer = null;
    function goB(n) {
      bi = (n + banners.length) % banners.length;
      bTrack.style.transform = 'translateX(' + (-bi * 100) + '%)';
      if (bDots) {
        $$('button', bDots).forEach(function (d, k) {
          d.classList.toggle('is-on', k === bi);
        });
      }
    }
    function startB() {
      if (bTimer || banners.length < 2) return;
      bTimer = setInterval(function () { goB(bi + 1); }, 5200);
    }
    function stopB() { if (bTimer) { clearInterval(bTimer); bTimer = null; } }

    var prev = $('#bnPrev'), next = $('#bnNext');
    if (prev) prev.addEventListener('click', function () { goB(bi - 1); stopB(); startB(); });
    if (next) next.addEventListener('click', function () { goB(bi + 1); stopB(); startB(); });
    if (bDots) {
      bDots.addEventListener('click', function (e) {
        var d = e.target.closest('[data-bdot]');
        if (!d) return;
        goB(parseInt(d.getAttribute('data-bdot'), 10));
        stopB(); startB();
      });
    }
    var banner = $('#banner');
    if (banner) {
      banner.addEventListener('mouseenter', stopB);
      banner.addEventListener('mouseleave', startB);
    }
    goB(0);
    startB();
  }

  /* ---------------------------------------------------------------- 公告列表 */
  var npList = $('#npList'), npCount = $('#npCount'), npAll = $('#npAll');
  var wall = notices.filter(function (n) { return n.cat === 'notice'; });
  if (!wall.length) wall = notices;
  if (npList && wall.length) {
    npList.innerHTML = wall.map(function (n, i) {
      return '<button type="button" class="np-item' + (n.cat === 'brief' ? ' is-brief' : '') +
        '" data-np="' + i + '">' +
        '<time>' + esc(n.date.slice(5)) + ' ' + esc(n.time) + '</time>' +
        '<b>' + esc(n.title) + '</b></button>';
    }).join('');
    npList.addEventListener('click', function (e) {
      var b = e.target.closest('[data-np]');
      if (b) openNotice(wall[parseInt(b.getAttribute('data-np'), 10)]);
    });
  }
  if (npCount) npCount.textContent = wall.length + ' 条';
  if (npAll) npAll.addEventListener('click', openList);

  /* ---------------------------------------------------------------- 卡片 */
  function cardHTML(f) {
    var cover = f.cover
      ? '<div class="board-thumb"><img src="' + esc(f.cover) + '" alt="" loading="lazy"></div>'
      : '<div class="board-thumb is-big"><span class="big-txt">' +
        esc(f.short || f.title) + '</span></div>';
    var body =
      '<div class="board-meta"><h3>' + esc(f.title) + '</h3>' +
        '<p>' + esc(f.desc || '') + '</p>' +
        '<div class="board-foot">' +
          '<span class="group-tag' + (f.url ? '' : ' tg-off') + '">' + esc(f.tag || '入口') + '</span>' +
          (f.url ? '<span class="open-link">打开 →</span>'
                 : '<span class="open-link muted">需登录</span>') +
        '</div>' +
      '</div>';
    var cls = 'board-card' + (f.cover ? '' : ' has-big') + (f.url ? '' : ' is-off');
    return f.url
      ? '<a class="' + cls + '" href="' + esc(f.url) + '" target="_blank" rel="noopener">' + cover + body + '</a>'
      : '<div class="' + cls + '">' + cover + body + '</div>';
  }
  function toolHTML(f) {
    var inner = '<span class="tc-ic">' + esc(f.icon || '▦') + '</span>' +
      '<span class="tc-txt"><b>' + esc(f.title) + '</b>' +
      '<span>' + esc(f.desc || '') + '</span></span>';
    return f.url
      ? '<a class="tool-card" href="' + esc(f.url) + '" target="_blank" rel="noopener">' + inner + '</a>'
      : '<div class="tool-card round" title="需要登录或暂未接入">' + inner + '</div>';
  }

  var host = $('#portalSections');
  if (host) {
    if (!cats.length) {
      host.innerHTML = '<div style="padding:40px 0;text-align:center;color:var(--muted)">' +
        '<div style="font-size:28px;margin-bottom:10px">▧</div>' +
        '<div>分区数据待补充 · 填入 site.config.js 的 portal 即可</div></div>';
    } else {
      host.innerHTML = cats.map(function (c, i) {
        var isTool = c.layout === 'tool';
        var grid = c.files.map(function (f) { return isTool ? toolHTML(f) : cardHTML(f); }).join('');
        return '<section class="portal-sec" id="sec' + i + '">' +
          '<div class="sec-head"><span class="sec-no">' + pad(i + 1) + '</span>' +
            '<h2>' + esc(c.icon) + ' ' + esc(c.name) + '</h2>' +
            '<span class="sec-count">' + c.files.length + ' 项</span></div>' +
          '<div class="' + (isTool ? 'tool-grid' : 'board-grid') + '">' + grid + '</div>' +
          '</section>';
      }).join('');
    }
  }

  /* ---------------------------------------------------------------- 杭州分公司动态 */
  var pnGrid = $('#branchNewsGrid');
  var branchNews = (S.news || []).filter(function (n) {
    return n.pinned || (n.tag || '').indexOf('杭州') >= 0;
  });
  if (!branchNews.length) branchNews = (S.news || []).slice(0, 3);
  if (pnGrid) {
    pnGrid.innerHTML = branchNews.slice(0, 3).map(function (n) {
      return '<a class="pn-card" href="' + esc(n.href || '#') + '" target="_blank" rel="noopener">' +
        '<div class="pn-thumb">' +
          (n.image ? '<img src="' + esc(n.image) + '" alt="" loading="lazy">' : '') +
          '<span class="pn-tag">' + esc(n.tag || '动态') + '</span>' +
        '</div>' +
        '<div class="pn-body">' +
          '<span class="pn-date">' + esc(n.date) + '</span>' +
          '<h3>' + esc(n.title) + '</h3>' +
          '<p>' + esc(n.desc || '') + '</p>' +
          '<span class="pn-more">阅读全文 <i>→</i></span>' +
        '</div></a>';
    }).join('');
  }

  /* ---------------------------------------------------------------- 滚动淡入（#074） */
  // 不用 IntersectionObserver：部分嵌入环境不触发；用位置轮询，行为可预期
  function setupReveal() {
    var blocks = $$('.pnews, .portal-sec');
    if (!blocks.length) return;
    var reduce = window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches;
    function show(el) {
      if (el.classList.contains('is-in')) return;
      el.classList.add('is-in');
      if (el.id === 'branchNews') {
        $$('.pn-card', el).forEach(function (c, i) {
          setTimeout(function () { c.classList.add('is-in'); }, 80 + i * 75);
        });
      }
    }
    if (reduce) { blocks.forEach(show); return; }
    function tick() {
      var vh = window.innerHeight || 800;
      blocks.forEach(function (b) {
        if (b.classList.contains('is-in')) return;
        var r = b.getBoundingClientRect();
        if (r.top < vh * 0.92 && r.bottom > 0) show(b);
      });
    }
    tick();
    window.addEventListener('scroll', tick, { passive: true });
    window.addEventListener('resize', tick);
    setInterval(tick, 400);          // 兜底：滚动事件被节流时仍能触发
  }

  /* ---------------------------------------------------------------- 侧栏导航 */
  var nav = $('#sideNav');
  if (nav) {
    nav.innerHTML =
      '<button type="button" class="nv on" data-jump="-1">' +
        '<span class="nv-ic">▦</span><span class="nv-tx">全部入口</span>' +
        '<span class="nv-n">' + totalCount + '</span></button>' +
      '<button type="button" class="nv" data-jump="news">' +
        '<span class="nv-ic">▤</span><span class="nv-tx">分公司动态</span>' +
        '<span class="nv-n">' + Math.min(3, (S.news || []).length) + '</span></button>' +
      cats.map(function (c, i) {
        return '<button type="button" class="nv" data-jump="' + i + '">' +
          '<span class="nv-ic">' + esc(c.icon) + '</span>' +
          '<span class="nv-tx">' + esc(c.name) + '</span>' +
          '<span class="nv-n">' + c.files.length + '</span></button>';
      }).join('');

    nav.addEventListener('click', function (e) {
      var b = e.target.closest('[data-jump]');
      if (!b) return;
      var raw = b.getAttribute('data-jump');
      var k = raw === 'news' ? 'news' : parseInt(raw, 10);
      lockUntil = Date.now() + 700;
      $$('.nv', nav).forEach(function (x) {
        x.classList.toggle('on', x.getAttribute('data-jump') === String(k));
      });
      if (k === -1) {
        window.scrollTo(0, 0);          // html{scroll-behavior:smooth} 负责过渡
      } else {
        var el = document.getElementById(k === 'news' ? 'branchNews' : 'sec' + k);
        if (el) {
          // 用 scrollTo 而非 scrollIntoView：iframe 嵌入时后者在部分环境不触发
          var top = el.getBoundingClientRect().top + window.scrollY - 16;
          window.scrollTo(0, top);
        }
      }
    });

    var lockUntil = 0;   // 点击跳转后的平滑滚动期间锁住，避免高亮来回跳
    function syncNav() {
      if (Date.now() < lockUntil) return;
      var y = window.scrollY + 150, cur = -1;
      var newsEl = document.getElementById('branchNews');
      if (newsEl && newsEl.offsetTop <= y) cur = 'news';
      cats.forEach(function (_, i) {
        var s = document.getElementById('sec' + i);
        if (s && s.offsetTop <= y) cur = i;
      });
      $$('.nv', nav).forEach(function (b) {
        b.classList.toggle('on', b.getAttribute('data-jump') === String(cur));
      });
    }
    var lastTick = 0;
    window.addEventListener('scroll', function () {
      var now = Date.now();
      if (now - lastTick < 70) return;
      lastTick = now;
      syncNav();
    }, { passive: true });
    syncNav();
    // 兜底：部分环境（后台标签页、嵌入 iframe）scroll 事件会被节流，定时校一次
    setInterval(function () { if (!document.hidden) syncNav(); }, 500);
  }

  /* ---------------------------------------------------------------- 搜索 */
  var input = $('#boardSearch'), count = $('#boardCount');
  function applyFilter() {
    var kw = (input ? input.value : '').trim().toLowerCase();
    var shown = 0;
    cats.forEach(function (_, i) {
      var sec = document.getElementById('sec' + i);
      if (!sec) return;
      var items = $$('.board-card, .tool-card', sec);
      var hit = 0;
      items.forEach(function (el) {
        var ok = !kw || el.textContent.toLowerCase().indexOf(kw) >= 0;
        el.style.display = ok ? '' : 'none';
        if (ok) hit++;
      });
      sec.style.display = hit ? '' : 'none';
      var cnt = $('.sec-count', sec);
      if (cnt) cnt.textContent = kw ? (hit + ' / ' + items.length + ' 项') : (items.length + ' 项');
      shown += hit;
    });
    if (count) {
      count.textContent = kw
        ? ('匹配 ' + shown + ' 项 · 共 ' + totalCount + ' 个入口')
        : ('共 ' + totalCount + ' 个入口 · 在线看板 ' + onlineCount + ' 个');
    }
  }
  if (input) input.addEventListener('input', applyFilter);
  applyFilter();

  setupReveal();

  /* ---------------------------------------------------------------- 搜索快捷键 */
  document.addEventListener('keydown', function (e) {
    if (e.key === 'k' && (e.metaKey || e.ctrlKey)) {
      e.preventDefault();
      if (input) input.focus();
    }
  });
})();
