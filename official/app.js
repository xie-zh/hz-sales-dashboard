/* ============================================================================
   组件实现 —— 严格按 Motion Prompt 的 Motion spec 移植到 CSS / JS
   ----------------------------------------------------------------------------
   // Motion Prompt #089 · MIT License (c) 2026 Shohei Ohtaki
   // Motion Prompt #075 · MIT License (c) 2026 Shohei Ohtaki
   // Motion Prompt #067 · MIT License (c) 2026 Shohei Ohtaki
   // Motion Prompt #074 · MIT License (c) 2026 Shohei Ohtaki
   // Motion Prompt #081 · MIT License (c) 2026 Shohei Ohtaki
   // Motion Prompt #086 · MIT License (c) 2026 Shohei Ohtaki
   // Motion Prompt #073 · MIT License (c) 2026 Shohei Ohtaki
   // Motion Prompt #079 · MIT License (c) 2026 Shohei Ohtaki
   // Motion Prompt #145 · MIT License (c) 2026 Shohei Ohtaki
   // Motion Prompt #147 · MIT License (c) 2026 Shohei Ohtaki
   // Motion Prompt #150 · MIT License (c) 2026 Shohei Ohtaki
   // Motion Prompt #069 · MIT License (c) 2026 Shohei Ohtaki
   // Motion Prompt #072 · MIT License (c) 2026 Shohei Ohtaki
   // Motion Prompt #057 · MIT License (c) 2026 Shohei Ohtaki
   // Motion Prompt #058 · MIT License (c) 2026 Shohei Ohtaki
   // Motion Prompt #071 · MIT License (c) 2026 Shohei Ohtaki
   // Motion Prompt #056 · MIT License (c) 2026 Shohei Ohtaki
   // Motion Prompt #084 · MIT License (c) 2026 Shohei Ohtaki
   // Motion Prompt #092 · MIT License (c) 2026 Shohei Ohtaki
   // Motion Prompt #153 · MIT License (c) 2026 Shohei Ohtaki
   // Motion Prompt #154 · MIT License (c) 2026 Shohei Ohtaki
   // https://showheyohtaki.github.io/motion-prompt/
   ----------------------------------------------------------------------------
   规格来源：本会话内 `motion-prompt.mjs get` 的输出；kit: ui / lp / lp2。
   网页项目按 SKILL.md「Web pages」一节，用项目自身的 CSS/JS 重建，
   保留时长、缓动与幅度，替换内容、配色与图片。
   ============================================================================ */
(function () {
  'use strict';
  var S = window.SITE || {};
  var reduced = window.matchMedia('(prefers-reduced-motion: reduce)').matches;
  var $ = function (s, r) { return (r || document).querySelector(s); };
  var $$ = function (s, r) { return Array.prototype.slice.call((r || document).querySelectorAll(s)); };
  var esc = function (v) {
    return String(v == null ? '' : v).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  };

  /* ------------------------------------------------------------- 登录会话 */
  // 员工通道为登录制：登录态存 localStorage，跨刷新保留；页面加载即恢复 UI
  var AUTH_KEY = 'dgg_staff_user';
  function getStaffUser() {
    try { return localStorage.getItem(AUTH_KEY) || ''; } catch (e) { return ''; }
  }
  function setStaffUser(u) {
    try { u ? localStorage.setItem(AUTH_KEY, u) : localStorage.removeItem(AUTH_KEY); } catch (e) {}
  }
  // 把所有「登录」按钮同步为当前会话状态：未登录显示「登录」，已登录显示绿点 + 用户名（点击退出）
  function authUI() {
    var user = getStaffUser();
    $$('[data-login]').forEach(function (btn) {
      if (user) {
        btn.textContent = user;
        btn.classList.add('is-authed');
        btn.setAttribute('title', '已登录：' + user + '（点击退出登录）');
      } else {
        btn.textContent = '登录';
        btn.classList.remove('is-authed');
        btn.removeAttribute('title');
      }
    });
  }
  // 员工通道页守卫：未登录直接访问 → 弹回对应首页并提示登录
  (function staffGuard() {
    var p = location.pathname;
    var isStaffPage = /final-staff\.html$/.test(p) || /\/staff\.html$/.test(p);
    if (!isStaffPage) return;
    if (!getStaffUser()) {
      var back = /final-staff\.html$/.test(p) ? 'final.html' : 'home.html';
      location.replace(back + '?needlogin=1');
      document.documentElement.style.display = 'none'; // 跳转前不渲染内容
    }
  })();


  /* ---------------------------------------------------------------- 主题 */
  var THEMES = {
    a: { file: 'theme-a.css', label: 'Apple 极简白' },
    b: { file: 'theme-b.css', label: 'Stripe / Linear 深色科技' },
    c: { file: 'theme-c.css', label: 'Microsoft / IBM 企业蓝' }
  };
  var q = new URLSearchParams(location.search);
  // final.html 用 FORCE_THEME 锁定 A 版；普通入口仍可切三版对比
  var theme = (window.FORCE_THEME || q.get('t') ||
               localStorage.getItem('site-theme') || 'a').toLowerCase();
  if (!THEMES[theme]) theme = 'a';
  if (!window.FORCE_THEME) localStorage.setItem('site-theme', theme);
  var link = document.createElement('link');
  link.rel = 'stylesheet'; link.href = THEMES[theme].file;
  document.head.appendChild(link);
  document.documentElement.setAttribute('data-theme', theme);
  window.SITE_THEME = theme;

  /* ------------------------------------------------ #069 骨架屏 / #072 空状态 */
  // #069: 灰块形状对应内容，每 0.9s 从左向右扫一道光
  // #072: 线描图标 + 「暂无内容」+ 说明，0.35s 淡入
  function skeleton(w, h, label) {
    var H = String(h || '160px');
    // 百分比高度需要父级有确定高度，否则 .skeleton 会塌陷
    var blockStyle = /%$/.test(H) ? ' style="height:100%"' : '';
    return '<div class="ph-block"' + blockStyle + '><div class="skeleton" style="width:' +
      (w || '100%') + ';height:' + H + '"></div>' +
      (label ? '<div class="ph-hint">' + esc(label) + '</div>' : '') + '</div>';
  }
  function emptyState(title, desc, icon) {
    return '<div class="empty"><div class="empty-ic">' + (icon || '▧') + '</div>' +
      '<div class="empty-t">' + esc(title || '内容待补充') + '</div>' +
      '<div class="empty-d">' + esc(desc || '在 _preview/site.config.js 中补入内容即可，框架位置已保留') + '</div>' +
      '</div>';
  }
  function textOrPlaceholder(v, fallback, cls) {
    if (v == null || v === '') {
      return '<span class="ph-text ' + (cls || '') + '">' + esc(fallback) + '</span>';
    }
    return esc(v);
  }

  /* ---------------------------------------------------------------- 内容注入 */
  function renderContent() {
    var brand = S.brand || {}, hero = S.hero || {};
    var brandName = brand.name || '公司名称待补充';
    var short = brand.shortName || brandName;

    // 页头 logo：缺失时显示文字占位块（保留尺寸，不塌陷）
    $$('[data-logo]').forEach(function (el) {
      el.innerHTML = brand.logo
        ? '<img src="' + esc(brand.logo) + '" alt="' + esc(short) + '">' +
          '<span class="logo-div"></span><span class="logo-txt">' + esc(short) + '</span>'
        : '<span class="logo-mark" aria-hidden="true"></span><span class="logo-txt">' + esc(short) + '</span>';
    });
    $$('[data-brand-name]').forEach(function (el) { el.textContent = brandName; });
    $$('[data-slogan]').forEach(function (el) {
      el.innerHTML = textOrPlaceholder(brand.slogan, '品牌标语待补充');
    });

    // 顶部公告条：无内容则隐藏，但保留 DOM
    var ann = $('#announcement');
    if (ann) {
      if (S.announcement) { ann.textContent = S.announcement; }
      else { ann.classList.add('is-empty'); }
    }

    // Hero
    // 标识条：把「杭州分公司」与「一站式企业服务平台」分层强调
    var eb = $('[data-hero-eyebrow]');
    if (eb) {
      var ebTxt = hero.eyebrow || '杭州分公司 · 一站式企业服务平台';
      var ebParts = String(ebTxt).split(/\s*-\s*/);
      eb.innerHTML = ebParts.length > 1
        ? '<b>' + esc(ebParts[0]) + '</b><i>·</i><span>' + esc(ebParts.slice(1).join(' - ')) + '</span>'
        : esc(ebTxt);
    }
    var ht = $('[data-hero-title]');
    if (ht) {
      if (hero.title) {
        var t = esc(hero.title);
        if (hero.highlight) {
          t = t.replace(esc(hero.highlight), '<span class="hl">' + esc(hero.highlight) + '</span>');
        }
        ht.innerHTML = t;
      } else {
        // 缺失时用完整示例文案撑起版式，不把「待补充」塞进标题里
        ht.innerHTML = '把数据变成<span class="hl">看得见的增长</span>';
      }
    }
    setText('[data-hero-sub]', hero.subtitle,
      '从业务数据到决策看板，让每一次判断都有据可依。');
    $$('[data-hero-primary]').forEach(function (el) { el.textContent = hero.primaryCta || '了解公司'; });
    $$('[data-hero-secondary]').forEach(function (el) { el.textContent = hero.secondaryCta || '了解公司'; });

    // Hero 主视觉：视频号置顶短片循环播放；无视频时显示封面 + 轻量占位说明
    var vc = S.videoChannel || {};
    var videoSrc = hero.video || vc.pinnedVideo;
    var hv = $('#heroVisual');
    if (hv) {
      if (videoSrc) {
        hv.innerHTML = '<video class="hero-video" autoplay loop muted playsinline preload="auto" ' +
          'poster="' + esc(hero.image || vc.poster || '') + '">' +
          '<source src="' + esc(videoSrc) + '" type="video/mp4"></video>';
        hv.classList.add('has-video');
        // 手机端自动播放兜底：iOS 低电量模式 / 部分浏览器需要用户手势才放行 autoplay，
        // 在 canplay、首次触摸/滚动/点击、页面回前台时反复尝试 play()（静默失败不报错）
        var vid = hv.querySelector('video');
        function kick() {
          if (!vid || vid.paused) {
            var pr = vid && vid.play();
            if (pr && pr.catch) pr.catch(function () {});
          }
        }
        if (vid) {
          vid.addEventListener('canplay', kick);
          vid.addEventListener('loadeddata', kick);
          ['touchstart', 'click', 'scroll'].forEach(function (ev) {
            window.addEventListener(ev, kick, { passive: true, once: true });
          });
          document.addEventListener('visibilitychange', function () {
            if (!document.hidden) kick();
          });
          setTimeout(kick, 400);
          setTimeout(kick, 1800);
        }
      } else if (hero.image) {
        var vTitle = vc.title || '置顶短片';
        hv.innerHTML = '<img src="' + esc(hero.image) + '" alt="">' +
          '<a class="video-slot" href="' + esc(vc.url || '#') + '" target="_blank" rel="noopener">' +
          (vc.cover ? '<img class="vs-thumb" src="' + esc(vc.cover) + '" alt="">' : '<span>▷</span>') +
          '<b>观看《' + esc(vTitle) + '》</b>' +
          (vc.duration ? '<span class="vs-dur">' + esc(vc.duration) + '</span>' : '') +
          '<em>视频号 ↗</em></a>';
      } else {
        hv.innerHTML = '<div class="hero-visual-ph">' +
          skeleton('100%', '100%', '主视觉待补充 · 建议 1600×900') +
          '<div class="hv-grid" aria-hidden="true"></div></div>';
      }
    }
    // 指标条：4 个槽位，缺数据出骨架
    var mb = $('#metricsRow');
    if (mb) {
      var ms = (S.metrics || []).slice(0, 4);
      var html = '';
      for (var i = 0; i < 4; i++) {
        var m = ms[i];
        html += '<div class="metric">' + (m
          ? '<b>' + esc(m.value) + '<i>' + esc(m.unit || '') + '</i></b><span>' + esc(m.label) + '</span>'
          : skeleton('62%', '46px') + skeleton('46%', '13px')) + '</div>';
      }
      mb.innerHTML = html;
    }

    // 关于我们 + 六大服务（合并板块）：左侧服务列表切换右侧大卡片，点卡片进详情
    var about = S.about || {};
    setText('[data-about-title]', about.title, '关于我们');
    setText('[data-about-body]', about.body, '公司背景、团队与使命介绍待补充。');
    $$('[data-about-address]').forEach(function (el) { el.textContent = about.address || ''; });
    $$('[data-about-note]').forEach(function (el) { el.textContent = about.addressNote || brandName; });

    var svcList = $('#svcList'), svcStage = $('#svcStage');
    if (svcList && svcStage) {
      var sols = S.solutions || [];
      var curSvc = 0;
      var two = function (n) { return ('0' + n).slice(-2); };

      function paintStage(i) {
        var x = sols[i] || {};
        var d = x.detail || {};
        var items = (d.items || []).map(function (t) { return '<li>' + esc(t) + '</li>'; }).join('');
        svcStage.innerHTML =
          '<article class="svc-card" data-solution="' + i + '" tabindex="0" role="button" ' +
            'aria-label="查看' + esc(x.title) + '服务详情">' +
            '<div class="sc-media">' +
              (x.image ? '<img src="' + esc(x.image) + '" alt="">' : skeleton('100%', '100%')) +
              '<span class="sc-no">' + two(i + 1) + '</span>' +
            '</div>' +
            '<div class="sc-body">' +
              '<span class="sc-kicker">服务 ' + two(i + 1) + ' / ' + two(sols.length) + '</span>' +
              '<h3>' + esc(x.title) + '</h3>' +
              '<p class="sc-desc">' + esc(x.desc || '') + '</p>' +
              (d.intro ? '<p class="sc-intro">' + esc(d.intro) + '</p>' : '') +
              (items ? '<ul class="sc-list">' + items + '</ul>' : '') +
              '<div class="sc-foot"><span class="btn btn-primary">查看服务详情</span>' +
                '<span class="sc-hint">点击卡片查看服务范围与咨询方式</span></div>' +
            '</div>' +
          '</article>';
        var card = $('.svc-card', svcStage);
        if (card) {
          card.classList.remove('is-in');
          void card.offsetWidth;          // 重排一次，让淡入动画重新播放
          card.classList.add('is-in');
        }
      }

      function paintList() {
        svcList.innerHTML = sols.map(function (x, i) {
          return '<button type="button" class="svc-item' + (i === curSvc ? ' is-on' : '') +
            '" data-svc="' + i + '">' +
            '<span class="si-ic">' + esc(x.icon || '◈') + '</span>' +
            '<span class="si-tx"><b>' + esc(x.title) + '</b><em>' + esc(x.short || x.desc || '') + '</em></span>' +
            '<span class="si-ar">›</span></button>';
        }).join('');
      }

      svcList.addEventListener('click', function (e) {
        var b = e.target.closest('[data-svc]');
        if (!b) return;
        var i = parseInt(b.getAttribute('data-svc'), 10);
        if (i === curSvc) return;
        curSvc = i;
        paintList();          // 整表重绘：逐项 toggle class 在部分环境下不会触发样式重算
        paintStage(i);
      });

      // 大卡片支持键盘触发详情
      svcStage.addEventListener('keydown', function (e) {
        if (e.key !== 'Enter' && e.key !== ' ') return;
        var c = e.target.closest && e.target.closest('.svc-card');
        if (!c) return;
        e.preventDefault();
        c.click();
      });

      paintList();
      paintStage(0);
    }

    // 发展历程（#145 滚动时间轴）
    var tl = $('#timelineList');
    if (tl) {
      var ts = S.timeline || [];
      tl.innerHTML = ts.length
        ? ts.map(function (t) {
            return '<li class="tl-item"><span class="tl-dot"></span><div class="tl-year">' + esc(t.year) +
              '</div><div class="tl-body"><b>' + esc(t.title) + '</b><p>' + esc(t.desc || '') + '</p></div></li>';
          }).join('')
        : [1, 2, 3, 4].map(function () {
            return '<li class="tl-item is-ph"><span class="tl-dot"></span>' +
              '<div class="tl-year ph-text">年份</div><div class="tl-body">' +
              skeleton('40%', '16px') + skeleton('82%', '12px') + '</div></li>';
          }).join('');
    }

    // 发展历程右侧：官网资质、服务承诺与集团数据
    var ta = $('#timelineAside');
    if (ta) {
      var a = S.timelineAside || {};
      var aItems = (a.items || []).map(function (t) { return '<li>' + esc(t) + '</li>'; }).join('');
      var aStats = (a.stats || []).map(function (t) {
        return '<div><b>' + esc(t.value) + '</b><span>' + esc(t.label) + '</span></div>';
      }).join('');
      ta.innerHTML =
        (a.image
          ? '<div class="ta-media"><img src="' + esc(a.image) + '" alt="集团办公环境" loading="lazy"></div>'
          : skeleton('100%', '220px', '配图待补充')) +
        '<div class="ta-body">' +
          (a.kicker ? '<span class="section-kicker">' + esc(a.kicker) + '</span>' : '') +
          '<h3>' + esc(a.title || '') + '</h3>' +
          (a.intro ? '<p class="ta-intro">' + esc(a.intro) + '</p>' : '') +
          (aItems ? '<ul class="ta-list">' + aItems + '</ul>' : '') +
          (aStats ? '<div class="ta-stats">' + aStats + '</div>' : '') +
          (a.note ? '<p class="ta-note">' + esc(a.note) + '</p>' : '') +
        '</div>';
    }

    // 客户案例（#058 轮播）
    var cs = $('#casesTrack');
    if (cs) {
      var cases = S.cases || [];
      cs.innerHTML = cases.length
        ? cases.map(function (c) {
            return '<div class="case-slide">' + (c.image ? '<img src="' + esc(c.image) + '" alt="">' : skeleton('100%', '190px')) +
              '<b>' + esc(c.name) + '</b><p>' + esc(c.desc || '') + '</p></div>';
          }).join('')
        : '<div class="grid-empty">' + emptyState('客户案例待补充', '填入 cases 后自动生成轮播', '◫') + '</div>';
    }

    // 新闻洞察：置顶文章排在最前，带缩略图与摘要
    var nb = $('#newsList');
    if (nb) {
      var news = (S.news || []).slice().sort(function (a, b) {
        return (b.pinned ? 1 : 0) - (a.pinned ? 1 : 0);
      });
      nb.innerHTML = news.length
        ? news.map(function (n) {
            return '<a class="news-item reveal' + (n.pinned ? ' is-pinned' : '') +
              '" href="' + esc(n.href || '#') + '" target="_blank" rel="noopener">' +
              '<div class="news-thumb">' + (n.image
                ? '<img src="' + esc(n.image) + '" alt="" loading="lazy">'
                : skeleton('100%', '100%')) + '</div>' +
              '<div class="news-text"><div class="news-meta">' +
                '<time>' + esc(n.date) + '</time>' +
                '<span class="news-badge">' + esc(n.tag || '动态') + '</span>' +
                (n.pinned ? '<span class="pin">置顶</span>' : '') +
              '</div><b>' + esc(n.title) + '</b><p>' + esc(n.desc || '') + '</p></div></a>';
          }).join('')
        : '<div class="grid-empty">' + emptyState('新闻洞察待补充', '填入 news 后自动渲染列表', '▤') + '</div>';
    }

    // 企业文化：核心理念 3 + 价值观 4 + 服务优势 6（文本取自官网「关于我们 /gywm」）
    var cl = $('#cultureList');
    if (cl) {
      var cu = S.culture || {};
      if (Array.isArray(cu)) cu = { core: cu };
      var core = cu.core || [], vals = cu.values || [], adv = cu.advantage || [];
      var IC = {
        flow: '<circle cx="6" cy="6" r="2.3"/><circle cx="18" cy="12" r="2.3"/><circle cx="6" cy="18" r="2.3"/>' +
              '<path d="M8.3 6H13a3 3 0 0 1 3 3v.7M8.3 18H13a3 3 0 0 0 3-3v-.7"/>',
        bolt: '<path d="M13.4 2.5 5 13.6h5.6l-.9 7.9 8.3-11.1h-5.5l.9-7.9Z"/>',
        target: '<circle cx="12" cy="12" r="8.6"/><circle cx="12" cy="12" r="4.1"/>' +
                '<circle cx="12" cy="12" r="1" fill="currentColor" stroke="none"/>',
        glass: '<path d="M2.2 12S5.9 5.6 12 5.6 21.8 12 21.8 12 18.1 18.4 12 18.4 2.2 12 2.2 12Z"/>' +
               '<circle cx="12" cy="12" r="2.8"/>',
        shield: '<path d="M12 3 5.2 5.9v6.2c0 4.2 2.9 7.6 6.8 9 3.9-1.4 6.8-4.8 6.8-9V5.9L12 3Z"/>' +
                '<path d="m9.2 11.9 2.2 2.2 3.6-3.9"/>',
        grid: '<rect x="3.4" y="3.4" width="7.2" height="7.2" rx="1.7"/>' +
              '<rect x="13.4" y="3.4" width="7.2" height="7.2" rx="1.7"/>' +
              '<rect x="3.4" y="13.4" width="7.2" height="7.2" rx="1.7"/>' +
              '<rect x="13.4" y="13.4" width="7.2" height="7.2" rx="1.7"/>'
      };
      var culIc = function (k) {
        return '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" ' +
          'stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">' +
          (IC[k] || IC.grid) + '</svg>';
      };

      var out = '';
      if (core.length) {
        out += '<div class="cul-core">' + core.map(function (c, i) {
          return '<article class="cul-card reveal" style="transition-delay:' + (i * 80) + 'ms">' +
            (c.no ? '<span class="cul-no">' + esc(c.no) + '</span>' : '') +
            (c.en ? '<span class="cul-en">' + esc(c.en) + '</span>' : '') +
            '<h3>' + esc(c.title) + '</h3><p>' + esc(c.desc || '') + '</p></article>';
        }).join('') + '</div>';
      }
      if (vals.length) {
        out += '<div class="cul-block"><div class="cul-sub reveal"><b>价值观</b>' +
          '<span>四条准则，落在每一次客户沟通里</span></div><div class="cul-values">' +
          vals.map(function (v, i) {
            return '<article class="cul-val reveal" style="transition-delay:' + (i * 70) + 'ms">' +
              '<i>' + esc(v.no || '') + '</i><b>' + esc(v.title) + '</b>' +
              '<p>' + esc(v.desc || '') + '</p></article>';
          }).join('') + '</div></div>';
      }
      if (adv.length) {
        out += '<div class="cul-block"><div class="cul-sub reveal"><b>服务优势</b>' +
          '<span>客户能感受到的六件事</span></div><div class="cul-adv">' +
          adv.map(function (a, i) {
            return '<article class="cul-adv-item reveal" style="transition-delay:' + (i * 55) + 'ms">' +
              '<div class="cul-ic">' + culIc(a.icon) + '</div><b>' + esc(a.title) + '</b>' +
              '<span>' + esc(a.desc || '') + '</span></article>';
          }).join('') + '</div></div>';
      }
      cl.innerHTML = out || '<div class="cul-core">' + [1, 2, 3].map(function () {
        return '<article class="cul-card is-ph"><h3 class="ph-text">标题</h3>' +
          skeleton('60%', '14px') + '</article>';
      }).join('') + '</div>';
    }

    // CTA / 页脚
    setText('[data-cta-title]', (S.cta || {}).title, '准备好开始了吗');
    setText('[data-cta-desc]', (S.cta || {}).desc, '补入一句行动引导文案，说明联系我们可以获得什么服务。');
    $$('[data-cta-btn]').forEach(function (el) { el.textContent = (S.cta || {}).button || '联系我们'; });
    setText('[data-footer-note]', (S.footer || {}).note, '© 2026 · 公司名称与备案信息待补充');
    var fa = $('[data-footer-address]');
    if (fa) fa.textContent = (S.footer || {}).address || '';
    var fas = $('[data-footer-address-sub]');
    if (fas) fas.textContent = (S.footer || {}).addressSub || '';

    // 主导航
    var nav = $('#navLinks');
    if (nav) {
      nav.innerHTML = (S.nav || []).map(function (n, i) {
        return '<a class="nav-link" href="' + esc(n.href) + '" data-mega="' + i + '">' + esc(n.label) + '</a>';
      }).join('');
    }

    // 手机端抽屉导航内容（主项 + mega 子项分组）
    var nd = $('#ndNav');
    if (nd) {
      nd.innerHTML = (S.nav || []).map(function (n) {
        var subs = '';
        if (n.mega && n.mega.length) {
          var links = n.mega.map(function (g) {
            return (g.links || []).map(function (l) {
              return '<a href="' + esc(l[1]) + '">' + esc(l[0]) + '</a>';
            }).join('');
          }).join('');
          subs = '<div class="nd-sub">' + links + '</div>';
        }
        return '<div class="nd-group">' +
          '<a class="nd-link" href="' + esc(n.href) + '">' + esc(n.label) + ' <i>→</i></a>' +
          subs + '</div>';
      }).join('');
    }
  }

  /* --------------------------------------------- 手机端左侧抽屉导航 */
  function navDrawer() {
    var burger = $('#navBurger'), drawer = $('#navDrawer'), scrim = $('#navScrim');
    if (!burger || !drawer) return;
    function open() {
      drawer.classList.add('is-open');
      if (scrim) scrim.classList.add('is-open');
      burger.classList.add('is-open');
      burger.setAttribute('aria-expanded', 'true');
      drawer.setAttribute('aria-hidden', 'false');
      document.body.style.overflow = 'hidden';
    }
    function close() {
      drawer.classList.remove('is-open');
      if (scrim) scrim.classList.remove('is-open');
      burger.classList.remove('is-open');
      burger.setAttribute('aria-expanded', 'false');
      drawer.setAttribute('aria-hidden', 'true');
      document.body.style.overflow = '';
    }
    burger.addEventListener('click', function () {
      drawer.classList.contains('is-open') ? close() : open();
    });
    drawer.addEventListener('click', function (e) {
      // 点导航链接跳转后自动收起抽屉
      if (e.target.closest && e.target.closest('a')) close();
    });
    document.addEventListener('click', function (e) {
      if (e.target.closest && e.target.closest('[data-nav-close]')) close();
    });
    document.addEventListener('keydown', function (e) {
      if (e.key === 'Escape' && drawer.classList.contains('is-open')) close();
    });
  }

  function setText(sel, value, placeholder) {
    $$(sel).forEach(function (el) { el.innerHTML = textOrPlaceholder(value, placeholder); });
  }

  /* ------------------------------------------------------------ #089 预加载 */
  // 0→100% at 0.25–2.1s (cubic ease-out)；2.3–2.85s 整屏上移 168px 退场
  function preloader() {
    var el = $('#preloader');
    if (!el) return;
    // ?skip=1 用于静态截图 / 自动化校验；正常访问保留预加载动画
    if (reduced || q.get('skip') === '1') { el.remove(); return; }
    var num = $('#preloadNum'), bar = $('#preloadBar');
    // 进度条时长：0.1–0.3 秒随机（用户要求快速闪过）
    var total = 100 + Math.random() * 200;
    var t0 = performance.now(), done = false;
    function finish() {
      if (done) return;
      done = true;
      if (num) num.textContent = '100%';
      if (bar) bar.style.width = '100%';
      el.classList.add('is-out');           // 整屏上移 168px，时长见 CSS transition
      setTimeout(function () { el.remove(); }, 300);
    }
    // 兜底：某些无头浏览器冻结 requestAnimationFrame，用定时器保证一定退场
    setTimeout(finish, total + 260);
    function tick(now) {
      if (done) return;
      var p = Math.min(1, Math.max(0, (now - t0) / total));
      var eased = 1 - Math.pow(1 - p, 3);
      var v = Math.round(eased * 100);
      if (num) num.textContent = v + '%';
      if (bar) bar.style.width = v + '%';
      if (p < 1) requestAnimationFrame(tick);
      else finish();
    }
    requestAnimationFrame(tick);
  }

  /* ------------------------------------------- #075 吸顶页头 / #079 进度条 */
  // #075: 24px 滚动内，12px 阴影透明度 0→0.16
  // #079: 顶部 6px 条，宽度 = 滚动比例
  function stickyHeader() {
    var head = $('#siteHeader'), bar = $('#progressBar');
    if (!head) return;
    function onScroll() {
      var y = window.scrollY || document.documentElement.scrollTop;
      var sh = Math.min(1, Math.max(0, y / 24));
      head.style.setProperty('--head-shadow', (0.16 * sh).toString());
      head.classList.toggle('is-stuck', y > 4);
      var h = document.documentElement.scrollHeight - window.innerHeight;
      if (bar) bar.style.width = (h > 0 ? (y / h * 100) : 0) + '%';
    }
    window.addEventListener('scroll', onScroll, { passive: true });
    onScroll();
  }

  /* ------------------------------------------------------ #067 巨型菜单 */
  // 展开 0→1 / 0.25s easeOutCubic，面板下落 12px；阴影 blur 18 / offset-y 8
  // 关闭 0.25s (1 − easeInOutCubic)
  function megaMenu() {
    var panel = $('#megaPanel'), inner = $('#megaInner');
    if (!panel) return;
    var hideTimer = null, current = -1;
    function show(i) {
      var n = (S.nav || [])[i];
      if (!n || !n.mega) return hide();
      current = i;
      var qrSrc = (S.contact || {}).qr || '';
      inner.innerHTML = n.mega.map(function (col) {
        return '<div class="mega-col"><h4>' + esc(col.title) + '</h4>' +
          (col.links || []).map(function (l) {
            return '<a href="' + esc(l[1]) + '">' + esc(l[0]) + '</a>';
          }).join('') + '</div>';
      }).join('') + '<div class="mega-visual">' + (qrSrc
        ? '<button type="button" class="mega-qr" data-contact aria-label="联系我们：扫码添加专属顾问微信">' +
            '<img src="' + esc(qrSrc) + '" alt="联系我们微信二维码">' +
            '<b>联系我们</b><span>扫码添加专属顾问</span>' +
          '</button>'
        : '<div class="mega-qr qr-empty"><span>微信二维码<br>待贴入</span><small>建议 240×240</small></div>') +
      '</div>';
      panel.classList.add('is-open');
      $$('.nav-link').forEach(function (a, k) { a.classList.toggle('is-active', k === i); });
    }
    function hide() {
      panel.classList.remove('is-open');
      $$('.nav-link').forEach(function (a) { a.classList.remove('is-active'); });
      current = -1;
    }
    // ?mega=N 自动展开时先「按住」，等真实鼠标移动后再交还给 hover 逻辑，
    // 否则无头截图/页面刚加载的杂散 mouseover 会在 120ms 后把它关掉。
    var holdMega = q.get('mega') !== null;
    if (holdMega) document.addEventListener('mousemove', function () { holdMega = false; }, { once: true });

    document.addEventListener('mouseover', function (e) {
      if (holdMega) return;
      var a = e.target.closest ? e.target.closest('.nav-link') : null;
      clearTimeout(hideTimer);
      if (a) { show(parseInt(a.dataset.mega, 10)); }
      else if (!e.target.closest('#megaPanel')) {
        hideTimer = setTimeout(hide, 120);
      }
    });
    panel.addEventListener('mouseleave', function () { hideTimer = setTimeout(hide, 120); });
    panel.addEventListener('mouseenter', function () { clearTimeout(hideTimer); });
    $('#siteHeader').addEventListener('mouseleave', function () { hideTimer = setTimeout(hide, 120); });

    // 调试/直达：?mega=0 直接展开对应主导航的巨型菜单
    var autoMega = q.get('mega');
    if (autoMega !== null) setTimeout(function () { show(parseInt(autoMega, 10) || 0); }, 200);
  }

  /* ------------------------------------------- #074 滚动淡入 / #145 时间轴 */
  // #074: 触发线 85% 视口高；0.7s cubic ease-out；28px 下→0；离开底部可重播
  // #145: 线随滚动延伸，节点越过 22px 后年份与事件淡入并右移 16px→0
  function revealOnScroll() {
    var items = $$('.reveal, .tl-item');
    if (reduced || q.get('all') === '1') {
      items.forEach(function (x) { x.classList.add('in'); });
      var tl = $('#timelineLine');
      if (tl) tl.style.height = '100%';
      return;
    }
    if ('IntersectionObserver' in window) {
      var io = new IntersectionObserver(function (entries) {
        entries.forEach(function (en) {
          if (en.isIntersecting) en.target.classList.add('in');
          else if (en.boundingClientRect.top > 0) en.target.classList.remove('in');
        });
      }, { rootMargin: '0px 0px -15% 0px', threshold: 0.05 });
      items.forEach(function (x) { io.observe(x); });
    } else {
      items.forEach(function (x) { x.classList.add('in'); });
    }
    // 时间轴：线条高度跟随视口中部触发线
    var line = $('#timelineLine'), list = $('#timelineList');
    if (line && list) {
      function upd() {
        var r = list.getBoundingClientRect();
        var trigger = window.innerHeight * 0.62;
        var p = Math.min(1, Math.max(0, (trigger - r.top) / Math.max(1, r.height)));
        line.style.height = (p * 100) + '%';
      }
      window.addEventListener('scroll', upd, { passive: true });
      upd();
    }
  }

  /* ---------------------------------- #081 悬停抬升 / #153 3D 倾斜 / #154 聚光灯 */
  // #081: 0.3s cubic ease-out，抬高 10px；阴影 blur 6→22、y 2→12、透明度 .08→.22
  // #153: 随光标最大倾斜 18°，反光跟随
  // #154: 只照亮光标周围约 120px
  function hoverEffects() {
    $$('.tilt-card').forEach(function (card) {
      card.addEventListener('mousemove', function (e) {
        if (reduced) return;
        var r = card.getBoundingClientRect();
        var px = (e.clientX - r.left) / r.width, py = (e.clientY - r.top) / r.height;
        card.style.transform = 'perspective(900px) rotateY(' + ((px - .5) * 36) + 'deg) rotateX(' + ((.5 - py) * 36) + 'deg)';
        card.style.setProperty('--mx', (px * 100) + '%');
        card.style.setProperty('--my', (py * 100) + '%');
      });
      card.addEventListener('mouseleave', function () { card.style.transform = ''; });
    });
    $$('.spotlight-card').forEach(function (card) {
      card.addEventListener('mousemove', function (e) {
        var r = card.getBoundingClientRect();
        card.style.setProperty('--mx', (e.clientX - r.left) + 'px');
        card.style.setProperty('--my', (e.clientY - r.top) + 'px');
        card.classList.add('is-lit');
      });
      card.addEventListener('mouseleave', function () { card.classList.remove('is-lit'); });
    });
  }

  /* ------------------------------------------------------------ #084 磁吸按钮 */
  // 按钮向光标方向移动两者距离的 0.3×，移开弹回
  function magneticButtons() {
    if (reduced) return;
    $$('.magnetic').forEach(function (btn) {
      btn.addEventListener('mousemove', function (e) {
        var r = btn.getBoundingClientRect();
        var dx = e.clientX - (r.left + r.width / 2), dy = e.clientY - (r.top + r.height / 2);
        btn.style.transform = 'translate(' + (dx * .3) + 'px,' + (dy * .3) + 'px)';
      });
      btn.addEventListener('mouseleave', function () { btn.style.transform = ''; });
    });
  }

  /* ------------------------------------------------------------ #073 视差滚动 */
  // 背景层按滚动 × 0.15 / 0.4 / 0.7 移动，内容 1.0
  function parallax() {
    if (reduced) return;
    function upd() {
      var y = window.scrollY || 0;
      $$('[data-parallax]').forEach(function (el) {
        var f = parseFloat(el.dataset.parallax) || 0.15;
        el.style.transform = 'translate3d(0,' + (-y * f) + 'px,0)';
      });
    }
    window.addEventListener('scroll', upd, { passive: true });
    upd();
  }

  /* ------------------------------------------------ #147 堆叠卡片（CSS sticky） */
  // sticky top 14 / 26 / 38px；下层缩 4.5%（两张 9%）并叠 16% 深色
  function stackingCards() {
    var cards = $$('.stack-card');
    if (!cards.length) return;
    function upd() {
      var n = cards.length;
      cards.forEach(function (c, i) {
        var r = c.getBoundingClientRect();
        var d = Math.min(1, Math.max(0, (14 + i * 12 - r.top) / 170));
        var under = Math.min(2, Math.max(0, n - 1 - i - (1 - d)));
        c.style.setProperty('--stack-scale', (1 - 0.045 * under).toFixed(3));
        c.style.setProperty('--stack-dim', (0.16 * under).toFixed(3));
      });
    }
    window.addEventListener('scroll', upd, { passive: true });
    upd();
  }

  /* ------------------------------------------------------------ #150 滚动监听 */
  // 当前章节高亮 0.25s cubic ease-out；触发线 = 视口顶部下 40px
  function scrollspy() {
    var links = $$('.toc a');
    if (!links.length) return;
    function upd() {
      var best = 0;
      links.forEach(function (a, i) {
        var t = document.querySelector(a.getAttribute('href'));
        if (t && t.getBoundingClientRect().top <= 40 + 40) best = i;
      });
      links.forEach(function (a, i) { a.classList.toggle('is-current', i === best); });
    }
    window.addEventListener('scroll', upd, { passive: true });
    upd();
  }

  /* ------------------------------------------------ #057 标签页 / #058 轮播 */
  // #057: 下划线 0.25s 滑动，内容 0.2s 叠化
  // #058: 每次 0.4s 切换，下方位置圆点
  function tabsAndCarousel() {
    var tabBar = $('#tabsBar');
    if (tabBar) {
      tabBar.addEventListener('click', function (e) {
        var b = e.target.closest('.tab-btn');
        if (!b) return;
        $$('.tab-btn').forEach(function (x) { x.classList.toggle('is-on', x === b); });
        var bar = $('#tabsUnderline');
        if (bar) { bar.style.left = b.offsetLeft + 'px'; bar.style.width = b.offsetWidth + 'px'; }
        $$('.tab-panel').forEach(function (p) {
          p.classList.toggle('is-on', p.dataset.panel === b.dataset.tab);
        });
      });
      var first = $('.tab-btn.is-on');
      if (first) requestAnimationFrame(function () {
        var bar = $('#tabsUnderline');
        if (bar) { bar.style.left = first.offsetLeft + 'px'; bar.style.width = first.offsetWidth + 'px'; }
      });
    }
    var track = $('#casesTrack'), idx = 0;
    var slides = $$('.case-slide', track || document);
    var dots = $('#casesDots');
    function go(n) {
      if (!slides.length || !track) return;
      idx = (n + slides.length) % slides.length;
      track.style.transform = 'translateX(-' + (idx * 100) + '%)';
      if (dots) $$('.dot', dots).forEach(function (d, i) { d.classList.toggle('is-on', i === idx); });
    }
    if (dots && slides.length) {
      dots.innerHTML = slides.map(function (_, i) { return '<button class="dot" data-i="' + i + '"></button>'; }).join('');
      dots.addEventListener('click', function (e) {
        var d = e.target.closest('.dot'); if (d) go(parseInt(d.dataset.i, 10));
      });
      go(0);
    }
    var prev = $('#casePrev'), next = $('#caseNext');
    if (prev) prev.addEventListener('click', function () { go(idx - 1); });
    if (next) next.addEventListener('click', function () { go(idx + 1); });
  }

  /* ------------------------------- #092 页面转场 + 员工通道（登录制，新标签打开） */
  // 点击员工通道：已登录 → 新标签打开（旧标签原地不动）；未登录 → 弹出登录框提示
  function staffEntry() {
    var curtain = $('#transitionCurtain');
    $$('[data-staff]').forEach(function (btn) {
      btn.addEventListener('click', function (e) {
        e.preventDefault();
        // 登录门槛：未登录只弹登录框，不跳转
        if (!getStaffUser()) {
          document.dispatchEvent(new CustomEvent('dgg:need-login', {
            detail: { err: '请先登录后查看员工通道', trigger: btn }
          }));
          return;
        }
        var url = window.FORCE_THEME ? 'final-staff.html' : ('staff.html?t=' + theme);
        // 注意：window.open 带 'noopener' 特性时按规范返回 null，
        // 会导致下方降级 location.href 把当前页也跳走（旧 bug），因此不带 noopener。
        var w = window.open(url, '_blank');
        if (w) { try { w.opener = null; } catch (err) {} }
        else { location.href = url; return; } // 仅弹窗被拦截时才原地打开
        if (curtain && !reduced) {
          curtain.classList.add('is-sweep');
          setTimeout(function () { curtain.classList.remove('is-sweep'); }, 700);
        }
      });
    });
  }

  /* ------------------------------------------------------- #056 登录轻提示 */
  // 登录按钮当前为占位：不改「免登录直达」的既定方案，点击给出明确反馈
  /* ------------------------------------------------------------------ 登录 / 注册 / 找回密码 */
  function loginEntry() {
    var modal = $('#loginModal'), body = $('#loginModalBody');
    if (!modal || !body) return;
    var lastFocus = null;
    // 演示账号：用户名与密码均为 admin（纯前端校验，源码可查）
    var CRED = { user: 'admin', pass: 'admin' };

    function toast(msg) {
      var host = $('#toastHost');
      if (!host) return;
      var t = document.createElement('div');
      t.className = 'toast';
      t.textContent = msg;
      host.appendChild(t);
      requestAnimationFrame(function () { t.classList.add('is-in'); });
      setTimeout(function () {
        t.classList.remove('is-in');
        setTimeout(function () { t.remove(); }, 360);
      }, 2200);
    }

    /* 关闭退场（Motion Prompt #049）：easeInOutCubic 0.24s 收拢淡出后由 JS 移除 is-open */
    var closing = false;
    function hide() {
      closing = false;
      modal.classList.remove('is-open', 'is-closing');
      modal.setAttribute('aria-hidden', 'true');
      document.body.style.overflow = '';
      if (lastFocus && lastFocus.focus) lastFocus.focus();
    }
    function close() {
      if (closing) return;
      if (reduced) { hide(); return; }
      closing = true;
      modal.classList.add('is-closing');
      setTimeout(hide, 240);
    }

    function show(err) {
      modal.classList.remove('is-closing');
      modal.classList.add('is-open');
      modal.setAttribute('aria-hidden', 'false');
      document.body.style.overflow = 'hidden';
      setTimeout(function () {
        var u = $('#lgUser', body);
        if (u && !err) u.focus();
      }, 60);
    }

    function shell(inner) {
      body.innerHTML = '<div class="lg-body">' + inner + '</div>';
      enter();
    }

    /* 新视图进场（Motion Prompt #057）：0.2s 淡入并上浮 8px，子元素依次错峰淡入 */
    function enter() {
      var fresh = body.firstElementChild;
      if (!fresh || reduced) return;
      fresh.classList.remove('lg-in');
      void fresh.offsetWidth; // 重置动画
      fresh.classList.add('lg-in');
      $$(':scope > *', fresh).forEach(function (el, k) {
        el.style.animationDelay = (k * 35) + 'ms';
      });
    }

    /* 视图切换（Motion Prompt #057 Tabs）：旧面板 0.2s 线性淡出 → 停 0.1s → 新面板淡入上浮 */
    var swapping = false;
    function swap(render) {
      var holder = body.firstElementChild;
      if (reduced || !holder || swapping) { render(); return; }
      swapping = true;
      holder.classList.add('lg-out');
      setTimeout(function () {
        swapping = false;
        render();
        var u = $('#lgUser', body);
        if (u) u.focus({ preventScroll: true });
      }, 300); // 0.2s 淡出 + 0.1s 停顿
    }

    /* ---------- 登录 ---------- */
    var EYE = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M1 12s4-7 11-7 11 7 11 7-4 7-11 7S1 12 1 12z"/><circle cx="12" cy="12" r="3"/></svg>';
    var EYE_OFF = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M17.94 17.94A10.07 10.07 0 0 1 12 20c-7 0-11-8-11-8a18.45 18.45 0 0 1 5.06-5.94M9.9 4.24A9.12 9.12 0 0 1 12 4c7 0 11 8 11 8a18.5 18.5 0 0 1-2.16 3.19m-6.72-1.07a3 3 0 1 1-4.24-4.24"/><line x1="1" y1="1" x2="23" y2="23"/></svg>';
    function viewLogin(err) {
      shell(
        '<span class="lg-kicker">Employee Sign in</span>' +
        '<h2 id="lgTitle">登录员工通道</h2>' +
        '<p class="lg-sub">使用公司分配的员工账号登录，进入数据工作台。</p>' +
        '<div class="lg-field"><label for="lgUser">用户名</label>' +
          '<input id="lgUser" type="text" autocomplete="username" placeholder="请输入用户名"><i class="lg-tick">✓</i></div>' +
        '<div class="lg-field has-eye"><label for="lgPass">密码</label>' +
          '<input id="lgPass" type="password" autocomplete="current-password" placeholder="请输入密码">' +
          '<button type="button" class="lg-eye" data-eye aria-label="显示密码" tabindex="-1">' + EYE + '</button>' +
          '<i class="lg-tick">✓</i></div>' +
        '<div class="lg-err"' + (err ? '' : ' hidden') + '>' + esc(err || '') + '</div>' +
        '<button type="button" class="btn btn-primary lg-submit" id="lgSubmit">登 录</button>' +
        '<div class="lg-links">' +
          '<button type="button" class="lg-link" data-lg="register">注册账号</button>' +
          '<span></span>' +
          '<button type="button" class="lg-link" data-lg="forgot">忘记密码</button>' +
        '</div>' +
        '<p class="lg-tip lg-tip-center">演示账号：admin / admin</p>');
    }

    /* ---------- 注册（占位） ---------- */
    function viewRegister() {
      shell(
        '<span class="lg-kicker">Sign up</span>' +
        '<h2 id="lgTitle">注册账号</h2>' +
        '<p class="lg-sub">自助注册暂未开放。可先通过微信或企大顺扫码登录，或联系管理员开通员工账号。</p>' +
        '<div class="lg-qrs">' +
          '<div class="lg-qr"><div><span>微信扫码</span><small>待接入</small></div><b>微信</b></div>' +
          '<div class="lg-qr"><div><span>企大顺扫码</span><small>待接入</small></div><b>企大顺</b></div>' +
        '</div>' +
        '<p class="lg-tip">扫码登录与企业账号体系尚未接入，此处为占位；接入后替换两张二维码即可。</p>' +
        '<button type="button" class="btn btn-ghost lg-submit" data-lg="login">返回登录</button>');
    }

    /* ---------- 忘记密码 ---------- */
    function viewForgot() {
      shell(
        '<span class="lg-kicker">Reset password</span>' +
        '<h2 id="lgTitle">联系集团 IT</h2>' +
        '<p class="lg-sub">密码重置由集团 IT 统一处理，请通过以下方式联系后重置。</p>' +
        '<ul class="lg-list">' +
          '<li><b>企业微信</b><span>搜索「集团 IT 服务台」发起会话</span></li>' +
          '<li><b>服务热线</b><span>400 666 5987 转 IT 支持</span></li>' +
          '<li><b>服务时间</b><span>工作日 9:00–18:00，一般 1 个工作日内处理</span></li>' +
        '</ul>' +
        '<p class="lg-tip">为保障账号安全，IT 会核对工号与所属部门后再重置密码。</p>' +
        '<button type="button" class="btn btn-primary lg-submit" data-lg="login">我知道了</button>');
    }

    /* ---------- 提交 ---------- */
    function submit() {
      var u = $('#lgUser', body), p = $('#lgPass', body);
      var user = ((u && u.value) || '').trim();
      var pass = (p && p.value) || '';
      if (!user || !pass) { setErr('请输入用户名和密码'); return; }
      if (user === CRED.user && pass === CRED.pass) {
        // 登录成功：写入会话（刷新/关浏览器后仍保留），刷新页头按钮为用户名，之后点员工通道即可进入
        setStaffUser(user);
        authUI();
        close();
        toast('登录成功，欢迎 ' + user);
      } else {
        setErr('用户名或密码错误，请重试');
        if (p) { p.value = ''; p.focus(); }
      }
    }
    function setErr(msg) {
      var box = $('.lg-err', body);
      if (!box) return;
      box.textContent = msg;
      box.removeAttribute('hidden');
      // 错误淡入（#057 内容进场节奏）+ 面板轻抖反馈
      box.classList.remove('is-in');
      void box.offsetWidth;
      box.classList.add('is-in');
      var panel = $('.lg-panel', modal);
      if (panel && !reduced) {
        panel.classList.remove('lg-shake');
        void panel.offsetWidth;
        panel.classList.add('lg-shake');
      }
      // 清空的密码字段同步撤掉小勾
      var pf = $('#lgPass', body);
      if (pf) { pf.value = ''; var ff = pf.closest('.lg-field'); if (ff) ff.classList.remove('is-filled'); }
    }

    // 输入反馈：字段填上内容后小勾弹入（focus 上浮 + label 变色由 CSS :focus-within 完成）
    body.addEventListener('input', function (e) {
      var f = e.target.closest && e.target.closest('.lg-field');
      if (f) f.classList.toggle('is-filled', !!String(e.target.value || '').trim());
    });

    modal.addEventListener('click', function (e) {
      var t = e.target;
      if (!t || !t.closest) return;
      var eye = t.closest('[data-eye]');
      if (eye) {
        var inp = eye.parentElement.querySelector('input');
        if (inp) {
          var showIt = inp.type === 'password';
          inp.type = showIt ? 'text' : 'password';
          eye.innerHTML = showIt ? EYE_OFF : EYE;
          eye.setAttribute('aria-label', showIt ? '隐藏密码' : '显示密码');
          inp.focus({ preventScroll: true });
        }
        return;
      }
      if (t.closest('[data-modal-close]')) { close(); return; }
      var nav = t.closest('[data-lg]');
      if (nav) {
        var v = nav.getAttribute('data-lg');
        // 视图切换走 #057 Tabs 节奏：淡出 → 停顿 → 新视图淡入上浮
        swap(function () {
          if (v === 'register') viewRegister();
          else if (v === 'forgot') viewForgot();
          else viewLogin();
        });
        return;
      }
      if (t.id === 'lgSubmit') submit();
    });
    modal.addEventListener('keydown', function (e) {
      if (e.key === 'Enter' && $('#lgSubmit', body)) { e.preventDefault(); submit(); }
    });
    document.addEventListener('keydown', function (e) {
      if (e.key === 'Escape' && modal.classList.contains('is-open')) close();
    });

    $$('[data-login]').forEach(function (btn) {
      btn.addEventListener('click', function (e) {
        e.preventDefault();
        // 已登录：点击 = 退出登录（绿点用户名按钮）
        if (getStaffUser()) {
          setStaffUser('');
          authUI();
          toast('已退出登录');
          return;
        }
        lastFocus = btn;
        viewLogin();
        show();
      });
    });

    // 员工通道在未登录时点击 → staffEntry 派发本事件：弹登录框并提示
    document.addEventListener('dgg:need-login', function (e) {
      var d = e.detail || {};
      lastFocus = d.trigger || null;
      viewLogin(d.err || '请先登录后查看员工通道');
      show();
    });

    // 调试直达：?login=1 / ?login=register / ?login=forgot；?needlogin=1 从员工通道页弹回
    var auto = q.get('login');
    if (auto !== null) {
      if (auto === 'register') viewRegister();
      else if (auto === 'forgot') viewForgot();
      else viewLogin();
      show();
    } else if (q.get('needlogin') !== null) {
      viewLogin('请先登录后查看员工通道');
      show();
    }

    // 页面加载即恢复会话：已登录用户页头显示用户名（绿点）
    authUI();
  }

  /* ------------------------------------------------------------ #071 徽标弹跳 */
  function badges() {
    $$('[data-badge]').forEach(function (b) {
      b.classList.add('badge-pop');
      setTimeout(function () { b.classList.remove('badge-pop'); }, 320);
    });
  }

  /* ------------------------------------------------------------------ 解决方案详情弹窗 */
  function solutionModal() {
    var modal = $('#solutionModal'), body = $('#solutionModalBody');
    if (!modal || !body) return;
    var lastFocus = null;

    function open(i) {
      var x = (S.solutions || [])[i] || {};
      var d = x.detail || {};
      var c = S.contact || {};
      var qr = d.qr || c.qr || '';
      var qrLabel = x.title ? esc(x.title) + ' · ' : '';
      var items = (d.items || []).map(function (t) { return '<li>' + esc(t) + '</li>'; }).join('');
      body.innerHTML =
        (d.image
          ? '<div class="md-hero"><img src="' + esc(d.image) + '" alt=""></div>'
          : '<div class="md-hero md-hero-ph">' + skeleton('100%', '100%') + '</div>') +
        '<div class="md-content">' +
          '<div class="md-scroll">' +
            '<span class="section-kicker">' + esc(x.icon || '◈') + ' 服务详情</span>' +
            '<h2 id="modalTitle">' + esc(x.title || '') + '</h2>' +
            '<p class="md-lead">' + esc(x.desc || '') + '</p>' +
            (d.intro ? '<p class="md-intro">' + esc(d.intro) + '</p>' : '') +
            (items ? '<h4>服务范围</h4><ul class="md-list">' + items + '</ul>' : '') +
          '</div>' +
          '<div class="md-contact">' +
            '<div class="md-contact-txt"><b>' + esc(c.title || '微信咨询') + '</b>' +
              '<span>' + esc(c.desc || '') + '</span>' +
              (c.phone ? '<em>咨询热线 ' + esc(c.phone) + '</em>' : '') +
            '</div>' +
            (qr
              ? '<div class="md-qr"><img src="' + esc(qr) + '" alt="微信二维码"></div>'
              : '<div class="md-qr qr-empty"><span>' + qrLabel + '微信二维码<br>待贴入</span><small>建议 240×240</small></div>') +
          '</div>' +
        '</div>';
      modal.classList.add('is-open');
      modal.setAttribute('aria-hidden', 'false');
      document.body.style.overflow = 'hidden';
      // 正文超出可视高度时才加底部渐隐，避免最后一行被硬切
      var sc = $('.md-scroll', modal);
      if (sc) sc.classList.toggle('is-overflowing', sc.scrollHeight - sc.clientHeight > 2);
      var btn = $('.modal-close', modal);
      if (btn) btn.focus();
    }

    function close() {
      modal.classList.remove('is-open');
      modal.setAttribute('aria-hidden', 'true');
      document.body.style.overflow = '';
      if (lastFocus && lastFocus.focus) lastFocus.focus();
    }

    document.addEventListener('click', function (e) {
      var t = e.target;
      if (!t || !t.closest) return;
      var card = t.closest('[data-solution]');
      if (card) {
        lastFocus = card;
        open(parseInt(card.getAttribute('data-solution'), 10) || 0);
        return;
      }
      if (t.closest('[data-modal-close]')) close();
    });
    document.addEventListener('keydown', function (e) {
      if (e.key === 'Escape' && modal.classList.contains('is-open')) close();
    });

    // 调试/直达：?open=0..5 可直接打开对应业务详情
    var autoOpen = q.get('open');
    if (autoOpen !== null) {
      setTimeout(function () { open(parseInt(autoOpen, 10) || 0); }, 260);
    }
  }

  /* ------------------------------------------------------------------ 公司人员介绍：左侧负责人列表 + 右侧介绍卡 + 成员弹窗 */
  function teamSection() {
    var list = (S.team || {}).members || [];
    var rail = $('#teamList'), card = $('#teamProfile');
    var modal = $('#teamModal'), mbody = $('#teamModalBody');
    if (!rail || !card) return;
    if (!list.length) {
      rail.innerHTML = '<div class="team-empty">团队数据待补充 · 填入 site.config.js 的 team.members 即可</div>';
      return;
    }
    var reduce = window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches;
    var cur = -1, lastFocus = null;

    function face(m, cls) {
      var src = m.photo || m.img;
      return src
        ? '<span class="' + cls + '"><img src="' + esc(src) + '" alt="' + esc(m.name) + '" loading="lazy"></span>'
        : '<span class="' + cls + ' is-empty">' + esc(String(m.name || '—').slice(0, 1)) + '</span>';
    }

    /* 左侧列表 */
    rail.innerHTML = list.map(function (m, i) {
      var n = (m.people || []).length;
      return '<button type="button" class="tl-item' + (i === 0 ? ' is-on' : '') + '" role="tab"' +
        ' aria-selected="' + (i === 0) + '" data-tl="' + i + '">' +
        face(m, 'tl-face') +
        '<span class="tl-txt"><b>' + esc(m.name) + '</b><span>' + esc(m.dept || m.title) + '</span></span>' +
        (n ? '<span class="tl-n">' + n + '</span>' : '') +
        '</button>';
    }).join('');

    /* 右侧介绍卡 */
    var swapSeq = 0; // 快速连点时只认最新一次切换
    function profile(i) {
      var m = list[i], ppl = m.people || [];
      cur = i;
      var stack = ppl.slice(0, 4).map(function (p) {
        return '<img src="' + esc(p.img) + '" alt="">';
      }).join('');
      // 总经理没有下属列表，用「事业部数 + 总人数」补齐底部信息
      var totalPeople = list.reduce(function (a, x) { return a + (x.people || []).length; }, 0);
      var headCount = list.filter(function (x) { return (x.people || []).length; }).length;
      var footInfo = ppl.length
        ? '团队 ' + ppl.length + ' 人'
        : headCount + ' 个事业部 · ' + totalPeople + ' 位同事';
      var html =
        (m.photo || m.img
          ? '<div class="tp-photo"><img src="' + esc(m.photo || m.img) + '" alt="' + esc(m.name) + '"></div>'
          : '<div class="tp-photo is-empty"><b>' + esc(String(m.name || '—').slice(0, 1)) +
            '</b><span>照片待补</span></div>') +
        '<div class="tp-body">' +
          '<span class="tp-kicker">' + (m.lead ? 'General Manager' : 'Business Unit Head') + '</span>' +
          '<h3>' + esc(m.name) + '</h3>' +
          '<span class="tp-role">' + esc(m.title || '') + '</span>' +
          (m.dept ? '<span class="tp-dept">' + esc(m.dept) + '</span>' : '') +
          '<p class="tp-bio">' + esc(m.desc || '') + '</p>' +
          '<div class="tp-foot">' +
            (stack ? '<span class="tp-stack">' + stack + '</span>' : '') +
            '<span class="tp-count">' + footInfo + '</span>' +
            (ppl.length ? '<span class="tp-cta">查看团队成员 <i>→</i></span>' : '') +
          '</div>' +
        '</div>';

      function apply() {
        card.className = 'team-profile' + (ppl.length ? '' : ' is-static');
        card.innerHTML = html;
        // 照片异步加载完成后再淡入，避免闪空白（封面照 loading="eager"）
        var im = card.querySelector('.tp-photo img');
        if (im && !reduce) {
          im.style.opacity = '0';
          im.style.transition = 'opacity .3s cubic-bezier(.22,.61,.36,1)';
          function fin() { im.style.opacity = '1'; }
          if (im.complete && im.naturalWidth) setTimeout(fin, 30);
          else im.addEventListener('load', fin, { once: true });
          im.addEventListener('error', fin, { once: true });
        }
      }

      // Motion Prompt #057（Tabs）：旧内容 0.18s 线性淡出 → 停 0.08s → 新内容 0.22s 淡入并上浮 8px
      swapSeq++;
      var my = swapSeq;
      if (reduce || !card.innerHTML) { apply(); return; }
      card.style.transition = 'opacity .18s linear, transform .18s linear';
      card.style.opacity = '0';
      card.style.transform = 'translateY(-4px)';
      setTimeout(function () {
        if (my !== swapSeq) return; // 期间又点了别人，交给最新请求
        apply();
        card.style.transition = 'none';
        card.style.opacity = '0';
        card.style.transform = 'translateY(8px)';
        requestAnimationFrame(function () {
          requestAnimationFrame(function () {
            if (my !== swapSeq) return;
            card.style.transition = 'opacity .22s cubic-bezier(.22,.61,.36,1), transform .22s cubic-bezier(.22,.61,.36,1)';
            card.style.opacity = '1';
            card.style.transform = 'none';
          });
        });
      }, 260);
    }

    function select(i) {
      if (i === cur) return;
      $$('.tl-item', rail).forEach(function (b) {
        var on = parseInt(b.getAttribute('data-tl'), 10) === i;
        b.classList.toggle('is-on', on);
        b.setAttribute('aria-selected', on ? 'true' : 'false');
      });
      profile(i);
      // 弹窗开着时点左侧列表 = 直接切换弹窗里的成员（与弹窗副标题承诺一致）
      if (modal && modal.classList.contains('is-open')) openModal(i, true);
    }

    rail.addEventListener('click', function (e) {
      var b = e.target.closest ? e.target.closest('[data-tl]') : null;
      if (b) select(parseInt(b.getAttribute('data-tl'), 10));
    });

    /* 成员弹窗（Motion Prompt #049：开 easeOutCubic 轻缩放进入；关 easeInOutCubic 收拢退场） */
    function modalHTML(i) {
      var m = list[i], ppl = m.people || [];
      var chips = list.map(function (x, k) {
        var n = (x.people || []).length;
        return '<button type="button" class="tmo-chip' + (k === i ? ' is-on' : '') +
          '" data-tm="' + k + '"' + (n ? '' : ' disabled') + ' aria-label="' + esc(x.name) + '">' +
          face(x, 'tmo-chip-face') + '<span>' + esc(x.name) + '</span></button>';
      }).join('');
      return '<div class="tmo-head"><h2 id="tmTitle">' + esc(m.name) + ' · ' + esc(m.dept) + '</h2>' +
          '<p>共 ' + ppl.length + ' 位成员</p>' +
          '<div class="tmo-chips" role="tablist">' + chips + '</div></div>' +
        '<div class="tmo-grid">' + ppl.map(function (p) {
          return '<div class="tmo-card">' +
            '<div class="tmo-photo">' +
              (p.dept ? '<span class="tmo-tag">' + esc(p.dept) + '</span>' : '') +
              '<img src="' + esc(p.img) + '" alt="' + esc(p.name) + '" loading="lazy"></div>' +
            '<b>' + esc(p.name) + '</b><span>' + esc(p.title || '') + '</span></div>';
        }).join('') + '</div>';
    }
    function staggerIn() {
      var cards = $$('.tmo-card', mbody);
      if (reduce) cards.forEach(function (c) { c.classList.add('is-in'); });
      else cards.forEach(function (c, k) { setTimeout(function () { c.classList.add('is-in'); }, 40 + k * 26); });
    }
    function openModal(i, isSwitch) {
      var m = list[i];
      if (!m || !(m.people || []).length || !modal || !mbody) return;
      if (isSwitch) {
        // 弹窗内切换负责人：内容 0.15s 淡出 → 换 → 淡入 + 重新 stagger（不动遮罩）
        mbody.style.transition = 'opacity .15s linear';
        mbody.style.opacity = '0';
        setTimeout(function () {
          mbody.innerHTML = modalHTML(i);
          mbody.style.opacity = '0';
          requestAnimationFrame(function () {
            requestAnimationFrame(function () {
              mbody.style.transition = 'opacity .2s cubic-bezier(.22,.61,.36,1)';
              mbody.style.opacity = '1';
            });
          });
          staggerIn();
        }, 150);
        return;
      }
      lastFocus = card;
      mbody.innerHTML = modalHTML(i);
      modal.classList.remove('is-closing');
      modal.classList.add('is-open');
      modal.setAttribute('aria-hidden', 'false');
      document.body.style.overflow = 'hidden';
      staggerIn();
      var btn = $('.modal-close', modal);
      if (btn) btn.focus();
    }
    var closing = false;
    function closeModal() {
      if (!modal || closing) return;
      if (reduce) { hideModal(); return; }
      closing = true;
      modal.classList.add('is-closing'); // CSS mdSink/mdFadeOut · 0.24s
      setTimeout(hideModal, 240);
    }
    function hideModal() {
      closing = false;
      modal.classList.remove('is-open', 'is-closing');
      modal.setAttribute('aria-hidden', 'true');
      document.body.style.overflow = '';
      if (lastFocus && lastFocus.focus) lastFocus.focus();
    }
    if (card) card.addEventListener('click', function () { openModal(cur); });
    if (modal) {
      document.addEventListener('click', function (e) {
        if (e.target.closest && e.target.closest('#teamModal [data-modal-close]')) closeModal();
        // 弹窗内点负责人小条 → 切换成员（同步左侧列表选中态）
        var chip = e.target.closest ? e.target.closest('#teamModal .tmo-chip[data-tm]') : null;
        if (chip && !chip.disabled) {
          var k = parseInt(chip.getAttribute('data-tm'), 10);
          if (k !== cur) select(k); // select 内部会带着弹窗一起切换
        }
      });
      document.addEventListener('keydown', function (e) {
        if (e.key === 'Escape' && modal.classList.contains('is-open')) closeModal();
      });
    }

    // 调试/直达：?team=1 直接选中并弹出该负责人的团队
    var auto = parseInt(q.get('team'), 10);
    var start = (auto >= 0 && auto < list.length) ? auto : 0;
    if (start > 0) {
      $$('.tl-item', rail).forEach(function (b) {
        var on = parseInt(b.getAttribute('data-tl'), 10) === start;
        b.classList.toggle('is-on', on);
        b.setAttribute('aria-selected', on ? 'true' : 'false');
      });
    }
    profile(start);
    if (auto >= 0 && (list[start].people || []).length) setTimeout(function () { openModal(start); }, 320);
  }

  /* ------------------------------------------------------------------ 联系我们（二维码）弹窗 */
  function contactModal() {
    var modal = $('#contactModal'), body = $('#contactModalBody');
    if (!modal || !body) return;
    var lastFocus = null;

    function open() {
      var c = S.contact || {}, a = S.about || {};
      body.innerHTML =
        '<div class="mc-body">' +
          '<span class="section-kicker">Contact us</span>' +
          '<h2 id="contactTitle">' + esc(c.title || '联系我们') + '</h2>' +
          '<p class="mc-desc">' + esc(c.desc || '扫码添加专属顾问，获取一对一服务方案') + '</p>' +
          (c.qr
            ? '<div class="mc-qr"><img src="' + esc(c.qr) + '" alt="微信二维码"></div>'
            : '<div class="mc-qr qr-empty"><span>微信二维码<br>待贴入</span><small>建议 240×240</small></div>') +
          (a.address ? '<p class="mc-addr"><b>' + esc(a.addressNote || '顶呱呱杭州分公司') + '</b><br>' + esc(a.address) + '</p>' : '') +
          (c.phone ? '<a class="mc-phone" href="tel:' + esc(String(c.phone).replace(/\s/g, '')) + '">咨询热线 ' + esc(c.phone) + '</a>' : '') +
        '</div>';
      modal.classList.add('is-open');
      modal.setAttribute('aria-hidden', 'false');
      document.body.style.overflow = 'hidden';
      var btn = $('.modal-close', modal);
      if (btn) btn.focus();
    }

    function close() {
      modal.classList.remove('is-open');
      modal.setAttribute('aria-hidden', 'true');
      document.body.style.overflow = '';
      if (lastFocus && lastFocus.focus) lastFocus.focus();
    }

    document.addEventListener('click', function (e) {
      var t = e.target;
      if (!t || !t.closest) return;
      var trigger = t.closest('[data-contact]');
      if (trigger) { lastFocus = trigger; open(); return; }
      if (t.closest('[data-modal-close]')) close();
    });
    document.addEventListener('keydown', function (e) {
      if (e.key === 'Escape' && modal.classList.contains('is-open')) close();
    });

    var autoOpen = q.get('contact');
    if (autoOpen !== null) setTimeout(open, 260);
  }

  /* ------------------------------------------- #092 页面过渡：离场幕布 + 入场淡入 */
  function pageTransition() {
    if (reduced || q.get('skip') === '1') return;
    var logo = (S.brand || {}).logo || '';

    // 入场：整页淡入（若脚本异常，1.2s 后强制兜底）
    var root = document.documentElement;
    root.classList.add('arriving');
    var arrived = false;
    function finishArrive() {
      if (arrived) return;
      arrived = true;
      root.classList.add('arrive-done');
      setTimeout(function () { root.classList.remove('arriving', 'arrive-done'); }, 700);
    }
    if (document.readyState === 'complete') setTimeout(finishArrive, 40);
    else window.addEventListener('load', function () { setTimeout(finishArrive, 40); });
    setTimeout(finishArrive, 1200);

    // 离场：站内 final.html / final-staff.html 之间用幕布横扫
    document.addEventListener('click', function (e) {
      if (e.button || e.metaKey || e.ctrlKey || e.shiftKey) return;
      var a = e.target.closest ? e.target.closest('a[href]') : null;
      if (!a || a.target === '_blank' || a.hasAttribute('download')) return;
      var href = a.getAttribute('href') || '';
      if (href.charAt(0) === '#' || !/^final(-staff)?\.html($|[?#])/.test(href)) return;
      e.preventDefault();
      var veil = document.createElement('div');
      veil.className = 'page-veil from-left is-cover';
      veil.innerHTML = logo ? '<img src="' + logo + '" alt="">' : '';
      document.body.appendChild(veil);
      requestAnimationFrame(function () {
        requestAnimationFrame(function () { veil.classList.remove('from-left'); });
      });
      setTimeout(function () { window.location.href = href; }, 470);
    }, true);
  }

  /* ------------------------------------------------------------------ 小青蛙咨询浮窗
     序列帧动画：4 帧 SVG（静立/眨眼/蹲起/腾空），跳一次 ≈ 0.85s 抛物线；
     每 6~9s 换一处边缘落脚点；气泡由 CSS 每 5s 循环（显示 3s）。点青蛙 = data-contact 弹微信二维码。 */
  function frogWidget() {
    var host = document.getElementById('frogWidget');
    var stage = host ? host.querySelector('.frog-stage') : null;
    if (!host || !stage) return;

    /* 序列帧（卡通青蛙：圆身 + 大眼 + 腮红） */
    var EYE_OPEN =
      '<circle cx="23" cy="21" r="8.6" fill="#fff"/><circle cx="41" cy="21" r="8.6" fill="#fff"/>' +
      '<circle cx="23" cy="22.5" r="3.9" fill="#25342a"/><circle cx="41" cy="22.5" r="3.9" fill="#25342a"/>' +
      '<circle cx="24.4" cy="21" r="1.3" fill="#fff"/><circle cx="42.4" cy="21" r="1.3" fill="#fff"/>';
    var EYE_BLINK =
      '<path d="M15.5 21 Q23 26.5 30.5 21" fill="none" stroke="#3d6b1e" stroke-width="2.6" stroke-linecap="round"/>' +
      '<path d="M33.5 21 Q41 26.5 48.5 21" fill="none" stroke="#3d6b1e" stroke-width="2.6" stroke-linecap="round"/>';
    function frog(body, eyes, blush, feet, mouth) {
      return '<svg viewBox="0 0 64 64" width="64" height="64" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">' +
        feet +
        '<ellipse cx="32" cy="' + body.cy + '" rx="' + body.rx + '" ry="' + body.ry + '" fill="#82ca3e" stroke="#5a8f27" stroke-width="2"/>' +
        '<ellipse cx="32" cy="' + (body.cy + 5) + '" rx="' + (body.rx * 0.55) + '" ry="' + (body.ry * 0.5) + '" fill="#dff3c0"/>' +
        eyes +
        blush +
        mouth +
        '</svg>';
    }
    var BLUSH = '<ellipse cx="15" cy="' + 41 + '" rx="4.4" ry="2.6" fill="#ffb3bc" opacity=".85"/>' +
      '<ellipse cx="49" cy="41" rx="4.4" ry="2.6" fill="#ffb3bc" opacity=".85"/>';
    var FRAMES = [
      /* 0 静立 */
      frog(
        { cy: 39, rx: 20, ry: 18 },
        EYE_OPEN, BLUSH,
        '<ellipse cx="20" cy="58" rx="7" ry="3.6" fill="#6fb534" stroke="#5a8f27" stroke-width="1.6"/>' +
        '<ellipse cx="44" cy="58" rx="7" ry="3.6" fill="#6fb534" stroke="#5a8f27" stroke-width="1.6"/>',
        '<path d="M25 44 Q32 50 39 44" fill="none" stroke="#3d6b1e" stroke-width="2.4" stroke-linecap="round"/>'
      ),
      /* 1 眨眼 */
      frog(
        { cy: 39, rx: 20, ry: 18 },
        EYE_BLINK, BLUSH,
        '<ellipse cx="20" cy="58" rx="7" ry="3.6" fill="#6fb534" stroke="#5a8f27" stroke-width="1.6"/>' +
        '<ellipse cx="44" cy="58" rx="7" ry="3.6" fill="#6fb534" stroke="#5a8f27" stroke-width="1.6"/>',
        '<path d="M25 44 Q32 50 39 44" fill="none" stroke="#3d6b1e" stroke-width="2.4" stroke-linecap="round"/>'
      ),
      /* 2 蹲起（蓄力，身体压扁） */
      frog(
        { cy: 43, rx: 23.5, ry: 14.5 },
        EYE_OPEN, BLUSH,
        '<ellipse cx="19" cy="56" rx="8" ry="3" fill="#6fb534" stroke="#5a8f27" stroke-width="1.6"/>' +
        '<ellipse cx="45" cy="56" rx="8" ry="3" fill="#6fb534" stroke="#5a8f27" stroke-width="1.6"/>',
        '<path d="M26 47 Q32 51.5 38 47" fill="none" stroke="#3d6b1e" stroke-width="2.4" stroke-linecap="round"/>'
      ),
      /* 3 腾空（拉伸，腿后蹬） */
      frog(
        { cy: 34, rx: 17.5, ry: 21 },
        EYE_OPEN, BLUSH,
        '<ellipse cx="19" cy="56.5" rx="5" ry="7.5" fill="#6fb534" stroke="#5a8f27" stroke-width="1.6"/>' +
        '<ellipse cx="45" cy="56.5" rx="5" ry="7.5" fill="#6fb534" stroke="#5a8f27" stroke-width="1.6"/>',
        '<path d="M25 40 Q32 46.5 39 40" fill="none" stroke="#3d6b1e" stroke-width="2.4" stroke-linecap="round"/>'
      )
    ];
    stage.innerHTML = FRAMES.map(function () { return '<div class="frog-frame"></div>'; }).join('');
    var cells = $$('.frog-frame', stage);
    cells.forEach(function (c, k) { c.innerHTML = FRAMES[k]; });

    var cur = -1; // 从 -1 起步，确保 show(0) 真正点亮第一帧
    function show(k) {
      if (k === cur) return;
      cells[cur] && cells[cur].classList.remove('is-on');
      cells[k] && cells[k].classList.add('is-on');
      cur = k;
    }
    show(0);

    /* 落脚点：视口左右两侧边缘的空白带（避开页头；钳制防止出界） */
    function clampPos(p) {
      var W = window.innerWidth, H = window.innerHeight;
      return {
        x: Math.min(Math.max(p.x, 8), Math.max(8, W - 86)),
        y: Math.min(Math.max(p.y, 130), Math.max(130, H - 150))
      };
    }
    function anchor() {
      var W = window.innerWidth, H = window.innerHeight;
      var x = Math.random() < 0.5
        ? 10 + Math.random() * 26
        : W - 78 - 10 - Math.random() * 26;
      var y = 130 + Math.random() * Math.max(160, H - 290);
      return clampPos({ x: x, y: y });
    }

    var pos = clampPos(anchor());
    if (window.innerWidth < 700) { pos = clampPos({ x: window.innerWidth - 78 - 8, y: window.innerHeight - 220 }); }
    host.style.left = pos.x + 'px';
    host.style.top = pos.y + 'px';

    if (reduced) return; // 降级：静止青蛙 + CSS 气泡照常循环

    /* 眨眼 */
    (function blink() {
      setTimeout(function () {
        show(1);
        setTimeout(function () { show(0); }, 160);
        blink();
      }, 2600 + Math.random() * 2200);
    })();

    /* 跳跃：抛物线位移 + 序列帧（蹲→腾空→落地蹲→静立） */
    var hopping = false;
    function hop() {
      var to = anchor();
      var from = clampPos({
        x: parseFloat(host.style.left) || pos.x,
        y: parseFloat(host.style.top) || pos.y
      });
      var t0 = null, DUR = 850, ARC = 70;
      hopping = true;
      show(2); // 蓄力
      setTimeout(function () {
        show(3); // 腾空
        requestAnimationFrame(function step(now) {
          if (t0 === null) t0 = now;
          var p = Math.min(1, (now - t0) / DUR);
          var e = p < .5 ? 2 * p * p : 1 - Math.pow(-2 * p + 2, 2) / 2; // easeInOutQuad
          var x = from.x + (to.x - from.x) * e;
          var y = from.y + (to.y - from.y) * e - Math.sin(Math.PI * p) * ARC;
          host.style.left = x + 'px';
          host.style.top = y + 'px';
          if (p < 1) requestAnimationFrame(step);
          else {
            show(2); // 落地缓冲
            setTimeout(function () { show(0); hopping = false; }, 140);
          }
        });
      }, 170);
    }
    (function schedule() {
      setTimeout(function () {
        if (!hopping) hop();
        schedule();
      }, 6000 + Math.random() * 3000);
    })();

    /* 点击青蛙：先蹦一下再弹二维码（气泡/弹窗由全局 data-contact 处理） */
    host.addEventListener('click', function () {
      if (hopping || reduced) return;
      var stage2 = host.querySelector('.frog-stage');
      if (stage2) {
        stage2.style.transition = 'transform .18s cubic-bezier(.34,1.56,.64,1)';
        stage2.style.transform = 'scale(1.14)';
        setTimeout(function () { stage2.style.transform = 'none'; }, 200);
      }
    });
  }

  /* ------------------------------------------------------------------ 启动 */
  function boot() {
    renderContent();
    navDrawer();
    preloader();
    stickyHeader();
    megaMenu();
    // 注意：teamSection 必须先于 revealOnScroll 渲染——左侧列表按钮复用了时间轴的
    // .tl-item 淡入类，若后渲染将永远拿不到 .in 而不可见（opacity:0）
    teamSection();
    revealOnScroll();
    hoverEffects();
    magneticButtons();
    parallax();
    stackingCards();
    scrollspy();
    tabsAndCarousel();
    staffEntry();
    loginEntry();
    badges();
    solutionModal();
    contactModal();
    frogWidget();
    pageTransition();
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot);
  else boot();
})();
