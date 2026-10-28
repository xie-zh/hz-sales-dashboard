/* ============================================================
 * MiniQA —— 数据问答助手（AI 开放问答 + 全工作空间数据注入）
 *  - 每个页面自动 fetch 同源 workspace_data.json（活动量+三力+人员+录音质检）
 *    作为 AI 回答背景；加载失败回退到页面内联数据。
 *  - 支持多轮对话（记住上下文）与流式生成（"思考"过程可见）。
 *  - 优先走服务端代理 SPARK_PROXY_URL 调用讯飞星火（密钥仅存于代理侧，不留前端）；
 *    代理未配置时回退到浏览器直连（仅过渡用，代理上线后应删除回退）。
 * ============================================================ */
(function () {
  "use strict";

  /* ---------- 服务端代理地址（部署 Cloudflare Worker 后填入） ---------- */
  /* 例：https://spark-proxy-xxxx.workers.dev  —— 密钥只在 Worker 侧，前端不再明文 */
  var SPARK_PROXY_URL = "";

  /* ---------- 直连降级配置（仅当 SPARK_PROXY_URL 为空时启用；代理上线后删除） ---------- */
  var SPARK = {
    appid: "735e0119",
    apiKey: "a0628d70545044f5bd0b00c9b49fa8a3",
    apiSecret: "ZDlmMjhlYzdmZWRiNzI4NWU5ZTdjOGU3",
    host: "spark-api.xf-yun.com",
    maxTokens: 1024,
    temperature: 0.4,
    models: [
      { name: "Spark Lite", path: "/v1.1/chat", domain: "lite" },
      { name: "Spark Max", path: "/v3.5/chat", domain: "generalv3.5" }
    ]
  };

  /* ---------- 数据读取（兼容 const 词法全局与 window 全局） ---------- */
  function getRaw() {
    try { if (typeof RAW !== "undefined" && RAW && RAW.total) return RAW; } catch (e) {}
    try { if (window.RAW && window.RAW.total) return window.RAW; } catch (e) {}
    return null;
  }
  function getDat() {
    try { if (typeof DATA !== "undefined" && DATA && DATA.agentSummary) return DATA; } catch (e) {}
    try { if (window.DATA && window.DATA.agentSummary) return window.DATA; } catch (e) {}
    return null;
  }
  function fmt(n) { return Math.round(Number(n) * 100) / 100; }
  function pct(n) { return (Math.round(Number(n) * 10) / 10) + "%"; }
  function pctf(n) { return pct(Number(n) * 100); }

  /* ---------- 统一数据背景：优先 fetch 同源 workspace_data.json ---------- */
  function buildContext() {
    return new Promise(function (resolve) {
      var fallback = function () {
        var raw = getRaw(), dat = getDat();
        if (raw && raw.total) return Promise.resolve(buildActivityText(raw));
        if (dat && dat.agentSummary) return Promise.resolve(buildCallText(dat));
        return Promise.resolve("");
      };
      if (typeof fetch !== "function") { return resolve(fallback()); }
      fetch("./workspace_data.json", { cache: "no-store" })
        .then(function (r) { return r.ok ? r.json() : Promise.reject(new Error("http " + r.status)); })
        .then(function (d) {
          var t = (d && d.context_text) || "";
          if (t && t.length > 30) resolve(t); else resolve(fallback());
        })
        .catch(function () { resolve(fallback()); });
    });
  }
  function buildActivityText(d) {
    var t = d.total || [], s = d.stats || [], meta = d.meta || {};
    var n = t.filter(function (x) { return x["是否达标"] === "达标"; }).length;
    var lines = ["【分公司活动量看板】数据时间：" + (meta.data_time || "未知")];
    lines.push("考核总人数" + fmt(sumN(s, "总人数")) + "人，" + s.length + "个部门；全员明细" + t.length + "人，完成率均值" + pctf(avgN(t, "完成率")) + "，达标" + n + "人（达标率" + pct(t.length ? n / t.length * 100 : 0) + "）");
    lines.push("见客合计" + fmt(sumN(t, "见客合计")) + "次，上门" + fmt(sumN(t, "当日上门量")) + "，外出" + fmt(sumN(t, "当日外出量")));
    return lines.join("\n");
  }
  function buildCallText(d) {
    var k = d.kpi || {}, rs = d.recStat || {}, ags = d.agentSummary || [], recs = d.recordings || [];
    var lines = ["【分公司销售通话录音质检报告】数据日期：" + ((d.meta && d.meta.date) || "未知")];
    lines.push("当日拨打" + fmt(k.call_total) + "通，接通" + fmt(k.connect_total) + "通（接通率" + pct((k.connect_rate || 0) * 100) + "），90天加微" + fmt(k.wechat_total) + "，月累计业绩¥" + fmt(k.profit_total));
    lines.push("录音库" + rs.total + "段，已转写质检" + recs.filter(function (r) { return r.transcribed; }).length + "段（≥10秒）");
    var sorted = ags.slice().sort(function (a, b) { return b.overall - a.overall; });
    if (sorted.length) lines.push("电邀规划师综合分：" + sorted.map(function (a) { return a.person + " " + fmt(a.overall) + "/5(样本" + a.count + ")"; }).join("、"));
    return lines.join("\n");
  }
  function sumN(a, k) { return (a || []).reduce(function (s, x) { return s + (Number(x[k]) || 0); }, 0); }
  function avgN(a, k) {
    var v = (a || []).map(function (x) { return Number(x[k]) || 0; }).filter(function (x) { return !isNaN(x); });
    return v.length ? v.reduce(function (p, c) { return p + c; }, 0) / v.length : 0;
  }

  /* ---------- 多轮对话上下文 ---------- */
  var conversation = []; // [{role:'user'|'assistant', content:'...'}]
  var MAX_HISTORY = 12;
  function pushTurn(role, content) {
    conversation.push({ role: role, content: content });
    if (conversation.length > MAX_HISTORY) conversation = conversation.slice(conversation.length - MAX_HISTORY);
  }

  /* ---------- 调星火：优先走代理（流式），否则浏览器直连 ---------- */
  function askLLM(question, context, onToken, onDone, onError) {
    if (SPARK_PROXY_URL && SPARK_PROXY_URL.length > 4) {
      return askViaProxy(question, context, onToken, onDone, onError);
    }
    return askDirect(question, context, onToken, onDone, onError);
  }

  /* ----- 服务端代理（密钥不出现在前端） ----- */
  function askViaProxy(question, context, onToken, onDone, onError) {
    var payload = {
      context: context || "",
      history: conversation.slice(0, -1), // 不含当前这轮 user
      question: question
    };
    fetch(SPARK_PROXY_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload)
    }).then(function (resp) {
      if (!resp.ok) { throw new Error("代理返回 " + resp.status); }
      if (!resp.body || !resp.body.getReader) { return resp.text().then(function (t) { onToken(t); onDone(); }); }
      var reader = resp.body.getReader();
      var decoder = new TextDecoder("utf-8");
      function pump() {
        return reader.read().then(function (r) {
          if (r.done) { onDone(); return; }
          onToken(decoder.decode(r.value, { stream: true }));
          return pump();
        });
      }
      return pump();
    }).catch(function (e) { onError(e); });
  }

  /* ----- 浏览器直连降级（过渡用，含密钥） ----- */
  function b64(buf) {
    var s = "";
    for (var i = 0; i < buf.length; i++) s += String.fromCharCode(buf[i]);
    return btoa(s);
  }
  function hmacSha256Base64(secret, message) {
    var enc = new TextEncoder();
    return crypto.subtle.importKey("raw", enc.encode(secret), { name: "HMAC", hash: "SHA-256" }, false, ["sign"])
      .then(function (key) { return crypto.subtle.sign("HMAC", key, enc.encode(message)); })
      .then(function (sig) { return b64(new Uint8Array(sig)); });
  }
  function sparkUrl(model) {
    var date = new Date().toUTCString();
    var signatureOrigin = "host: " + SPARK.host + "\ndate: " + date + "\nGET " + model.path + " HTTP/1.1";
    return hmacSha256Base64(SPARK.apiSecret, signatureOrigin).then(function (signature) {
      var authorizationOrigin = 'api_key="' + SPARK.apiKey + '", algorithm="hmac-sha256", headers="host date request-line", signature="' + signature + '"';
      var authorization = btoa(authorizationOrigin);
      return "wss://" + SPARK.host + model.path + "?authorization=" + encodeURIComponent(authorization) + "&date=" + encodeURIComponent(date) + "&host=" + encodeURIComponent(SPARK.host);
    });
  }
  function askDirect(question, context, onToken, onDone, onError) {
    var models = SPARK.models || [];
    var systemMsg = "你是「小析」，分公司业务数据助手。请用简体中文、分点简洁回答，优先引用【数据背景】中的真实数字，不要编造。最多200字。\n【数据背景】\n" + context;
    var messages = [{ role: "user", content: systemMsg + "\n【用户问题】\n" + question }];
    // 直连降级仅支持单轮（上下文已在 system 内）
    var tryOne = function (idx) {
      if (idx >= models.length) { onError(new Error("所有模型均不可用")); return; }
      var model = models[idx];
      sparkUrl(model).then(function (url) {
        var ws = new WebSocket(url);
        var done = false, failed = false, full = "";
        var timer = setTimeout(function () { if (!done) { done = true; try { ws.close(); } catch (e) {} tryOne(idx + 1); } }, 25000);
        ws.onopen = function () {
          ws.send(JSON.stringify({
            header: { app_id: SPARK.appid, uid: "miniqa" },
            parameter: { chat: { domain: model.domain, temperature: SPARK.temperature, max_tokens: SPARK.maxTokens } },
            payload: { message: { text: [{ role: "user", content: systemMsg }, { role: "user", content: question }] } }
          }));
        };
        ws.onmessage = function (ev) {
          var d; try { d = JSON.parse(ev.data); } catch (e) { return; }
          if (d.header && d.header.code !== 0) {
            if (!done) { done = true; clearTimeout(timer); failed = true; try { ws.close(); } catch (e) {} tryOne(idx + 1); }
            return;
          }
          var t = (d.payload && d.payload.choices && d.payload.choices.text) || [];
          t.forEach(function (x) { if (x && x.content) { full += x.content; onToken(x.content); } });
          if (d.header && d.header.status === 2) { if (!done) { done = true; clearTimeout(timer); onDone(); } }
        };
        ws.onerror = function () { if (!done && !failed) { done = true; clearTimeout(timer); tryOne(idx + 1); } };
      }).catch(function () { if (!done) { done = true; tryOne(idx + 1); } });
    };
    tryOne(0);
  }

  /* ============================================================
   * UI 渲染（自注入样式 + DOM + 对话流）
   * ============================================================ */
  function injectStyle() {
    var css = [
      "#qaFab{position:fixed;right:20px;bottom:20px;z-index:99999;width:52px;height:52px;border-radius:50%;border:none;cursor:pointer;",
      "  background:linear-gradient(135deg,#4f46e5,#7c3aed);color:#fff;font-size:24px;box-shadow:0 6px 18px rgba(79,70,229,.45);transition:transform .15s}",
      "#qaFab:hover{transform:scale(1.08)}",
      "#qaWidget{position:fixed;right:20px;bottom:84px;z-index:99999;width:380px;max-width:calc(100vw - 40px);height:520px;max-height:calc(100vh - 110px);",
      "  background:var(--card,#fff);color:var(--text,#1a1a2e);border:1px solid var(--border,#e5e7eb);border-radius:14px;box-shadow:0 12px 40px rgba(0,0,0,.22);",
      "  display:flex;flex-direction:column;overflow:hidden;font-family:var(--font,'Segoe UI','PingFang SC','Microsoft YaHei',sans-serif)}",
      "#qaWidget.qa-hidden{display:none}",
      "#qaHeader{display:flex;align-items:center;justify-content:space-between;padding:12px 14px;background:linear-gradient(135deg,#4f46e5,#7c3aed);color:#fff;font-weight:700;font-size:15px}",
      "#qaClose{background:none;border:none;color:#fff;font-size:20px;cursor:pointer;line-height:1}",
      "#qaClear{background:none;border:1px solid rgba(255,255,255,.5);color:#fff;font-size:11px;padding:2px 8px;border-radius:8px;cursor:pointer;margin-right:6px}",
      "#qaBadge{font-size:10px;background:rgba(255,255,255,.22);padding:2px 7px;border-radius:8px;margin-left:6px;font-weight:600}",
      "#qaBody{flex:1;overflow-y:auto;padding:12px;font-size:13.5px;line-height:1.6}",
      "#qaBody .qa-msg{margin-bottom:10px;padding:9px 12px;border-radius:10px;white-space:pre-wrap;word-break:break-word}",
      "#qaBody .qa-bot{background:var(--bg,#f4f6f9);color:var(--text,#1a1a2e)}",
      "#qaBody .qa-user{background:#4f46e5;color:#fff;margin-left:36px}",
      "#qaBody .qa-thinking{color:#7c3aed;font-style:italic}",
      "#qaBody .qa-cursor::after{content:'▋';animation:qablink 1s steps(2) infinite;color:#7c3aed}",
      "@keyframes qablink{0%,50%{opacity:1}51%,100%{opacity:0}}",
      "#qaChips{display:flex;flex-wrap:wrap;gap:6px;padding:0 12px 8px}",
      "#qaChips .qa-chip{font-size:12px;padding:5px 10px;border:1px solid var(--border,#e5e7eb);border-radius:14px;background:var(--bg,#f4f6f9);color:var(--muted,#6b7280);cursor:pointer}",
      "#qaChips .qa-chip:hover{border-color:#4f46e5;color:#4f46e5}",
      "#qaInputBar{display:flex;gap:8px;padding:10px 12px;border-top:1px solid var(--border,#e5e7eb)}",
      "#qaInput{flex:1;padding:9px 12px;border:1px solid var(--border,#e5e7eb);border-radius:9px;font-size:13.5px;outline:none;background:var(--bg,#f4f6f9);color:var(--text,#1a1a2e)}",
      "#qaInput:focus{border-color:#4f46e5}",
      "#qaSend{padding:0 16px;border:none;border-radius:9px;background:#4f46e5;color:#fff;font-weight:600;cursor:pointer;font-size:13.5px}",
      "#qaSend:hover{background:#4338ca}"
    ].join("\n");
    var st = document.createElement("style");
    st.textContent = css;
    document.head.appendChild(st);
  }
  function buildUI() {
    injectStyle();
    var fab = document.createElement("button");
    fab.id = "qaFab"; fab.textContent = "💬"; fab.title = "数据助手";
    var box = document.createElement("div");
    box.id = "qaWidget"; box.className = "qa-hidden";
    var chips = ["整体达标情况", "三力有效人力", "人员业绩概况", "哪位电邀规划师最好", "录音质检硬伤", "公司整体建议"];
    box.innerHTML =
      '<div id="qaHeader"><span>📊 数据助手 · 小析<span id="qaBadge">AI</span></span>' +
      '<span><button id="qaClear" title="清空对话上下文">清空</button><button id="qaClose">×</button></span></div>' +
      '<div id="qaBody"></div>' +
      '<div id="qaChips">' + chips.map(function (c) { return '<span class="qa-chip">' + c + "</span>"; }).join("") + "</div>" +
      '<div id="qaInputBar"><input id="qaInput" placeholder="问任意数据问题，如「公司整体情况？」" /><button id="qaSend">发送</button></div>';
    document.body.appendChild(fab);
    document.body.appendChild(box);
    var body = box.querySelector("#qaBody");
    body.appendChild(msg("bot", "你好，我是数据助手「小析」🤖\n我已掌握工作空间全部看板数据（活动量 / 员工三力 / 人员分析 / 通话录音质检），可连续对话、记得上下文。用自然语言问我即可，例如「整体达标情况？」「哪位电邀规划师最需要改进？」「公司整体建议？」。"));

    function msg(role, text, extraCls) {
      var d = document.createElement("div");
      d.className = "qa-msg " + (role === "user" ? "qa-user" : "qa-bot") + (extraCls ? " " + extraCls : "");
      d.textContent = text;
      return d;
    }
    var busy = false;
    function send(q) {
      q = (q || "").trim();
      if (!q || busy) return;
      busy = true;
      body.appendChild(msg("user", q));
      pushTurn("user", q);
      // 思考过程：两步可见
      var think = msg("bot", "📚 正在结合全部看板数据…", "qa-thinking qa-cursor");
      body.appendChild(think);
      body.scrollTop = body.scrollHeight;
      buildContext().then(function (ctx) {
        if (!ctx) {
          think.textContent = "⚠️ 当前页面工作空间数据尚未加载，请稍候刷新后再试。";
          think.classList.remove("qa-thinking", "qa-cursor");
          busy = false; return;
        }
        think.textContent = "🧠 小析正在分析…";
        // 流式：首字到达后进入"生成中"光标态
        var started = false;
        askLLM(q, ctx,
          function (token) {
            if (!started) { started = true; think.textContent = ""; think.classList.remove("qa-thinking"); think.classList.add("qa-cursor"); }
            think.textContent += token;
            body.scrollTop = body.scrollHeight;
          },
          function () {
            think.classList.remove("qa-cursor");
            pushTurn("assistant", think.textContent);
            busy = false;
          },
          function (e) {
            think.textContent = "⚠️ AI 服务暂不可用：" + (e && e.message ? e.message : e) +
              "\n\n可稍后重试；若持续失败，可能是代理/接口额度或网络限制。";
            think.classList.remove("qa-thinking", "qa-cursor");
            busy = false;
          });
      }).catch(function () {
        think.textContent = "⚠️ 数据背景加载失败，请刷新页面后重试。";
        think.classList.remove("qa-thinking", "qa-cursor");
        busy = false;
      });
    }
    fab.addEventListener("click", function () { box.classList.toggle("qa-hidden"); });
    box.querySelector("#qaClose").addEventListener("click", function () { box.classList.add("qa-hidden"); });
    box.querySelector("#qaClear").addEventListener("click", function () {
      conversation = [];
      body.innerHTML = "";
      body.appendChild(msg("bot", "🧹 已清空对话上下文，新话题开始。"));
    });
    box.querySelector("#qaSend").addEventListener("click", function () { send(box.querySelector("#qaInput").value); box.querySelector("#qaInput").value = ""; });
    box.querySelector("#qaInput").addEventListener("keydown", function (e) { if (e.key === "Enter") { send(this.value); this.value = ""; } });
    Array.prototype.forEach.call(box.querySelectorAll(".qa-chip"), function (c) {
      c.addEventListener("click", function () { send(c.textContent); });
    });
  }

  if (typeof document !== "undefined") {
    if (document.readyState === "loading") {
      document.addEventListener("DOMContentLoaded", buildUI);
    } else {
      buildUI();
    }
  }
  if (typeof globalThis !== "undefined") {
    globalThis.MiniQA = { askLLM: askLLM, buildContext: buildContext, version: "3.0" };
  }
})();
