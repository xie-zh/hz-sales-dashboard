/* ============================================================
 * spark_proxy.js  ——  讯飞星火服务端代理（Cloudflare Worker）
 * 作用：浏览器不再持有任何星火密钥，改为调用本 Worker；
 *       Worker 在服务端用密钥连接讯飞星火，把回答流式回传给前端。
 * 部署：wrangler deploy  （需先 `wrangler login`）
 * ============================================================ */
const SPARK = {
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

function b64(buf) {
  let s = "";
  const bytes = new Uint8Array(buf);
  for (let i = 0; i < bytes.length; i++) s += String.fromCharCode(bytes[i]);
  return btoa(s);
}
async function hmacSha256Base64(secret, message) {
  const enc = new TextEncoder();
  const key = await crypto.subtle.importKey("raw", enc.encode(secret), { name: "HMAC", hash: "SHA-256" }, false, ["sign"]);
  const sig = await crypto.subtle.sign("HMAC", key, enc.encode(message));
  return b64(sig);
}
async function sparkUrl(model) {
  const date = new Date().toUTCString();
  const signatureOrigin = "host: " + SPARK.host + "\ndate: " + date + "\nGET " + model.path + " HTTP/1.1";
  const signature = await hmacSha256Base64(SPARK.apiSecret, signatureOrigin);
  const authorizationOrigin = 'api_key="' + SPARK.apiKey + '", algorithm="hmac-sha256", headers="host date request-line", signature="' + signature + '"';
  const authorization = btoa(authorizationOrigin);
  return "wss://" + SPARK.host + model.path + "?authorization=" + encodeURIComponent(authorization) + "&date=" + encodeURIComponent(date) + "&host=" + encodeURIComponent(SPARK.host);
}

async function callSpark(model, messages) {
  const url = await sparkUrl(model);
  const resp = await fetch(url, { headers: { "Upgrade": "websocket" } });
  const ws = resp.webSocket;
  if (!ws) throw new Error("星火连接失败: " + model.name);
  await ws.accept();

  const stream = new ReadableStream({
    start(controller) {
      const enc = new TextEncoder();
      let closed = false;
      const close = (err) => {
        if (closed) return;
        closed = true;
        try { if (err) controller.error(err); else controller.close(); } catch (e) {}
        try { ws.close(); } catch (e) {}
      };
      ws.addEventListener("message", (ev) => {
        let d; try { d = JSON.parse(ev.data); } catch (e) { return; }
        if (d.header && d.header.code !== 0) { close(new Error("星火错误 " + d.header.code + " " + d.header.message)); return; }
        const arr = (d.payload && d.payload.choices && d.payload.choices.text) || [];
        arr.forEach((x) => { if (x && x.content) controller.enqueue(enc.encode(x.content)); });
        if (d.header && d.header.status === 2) close();
      });
      ws.addEventListener("error", () => close(new Error("星火连接中断")));
      ws.addEventListener("close", () => close());
      ws.send(JSON.stringify({
        header: { app_id: SPARK.appid, uid: "miniqa-proxy" },
        parameter: { chat: { domain: model.domain, temperature: SPARK.temperature, max_tokens: SPARK.maxTokens } },
        payload: { message: { text: messages } }
      }));
    }
  });
  return stream;
}

async function handlePost(request) {
  let body;
  try { body = await request.json(); } catch (e) { return new Response("bad json", { status: 400 }); }
  const context = (body.context || "").toString();
  const history = Array.isArray(body.history) ? body.history : [];
  const question = (body.question || "").toString();
  if (!question) return new Response("empty question", { status: 400 });

  const systemMsg = "你是「小析」，分公司业务数据助手。请用简体中文、分点简洁回答，优先引用【数据背景】中的真实数字，不要编造。最多200字。\n【数据背景】\n" + context;
  // 星火仅支持 user/assistant 角色：把系统指令+背景作为首条 user 消息
  const messages = [{ role: "user", content: systemMsg }];
  history.forEach((h) => {
    if (h && (h.role === "user" || h.role === "assistant") && h.content) messages.push({ role: h.role, content: h.content });
  });
  messages.push({ role: "user", content: question });

  let lastErr = null;
  for (const model of SPARK.models) {
    try {
      const stream = await callSpark(model, messages);
      return new Response(stream, {
        headers: {
          "Content-Type": "text/plain; charset=utf-8",
          "Cache-Control": "no-store",
          "Access-Control-Allow-Origin": "*",
          "Access-Control-Allow-Headers": "Content-Type",
          "Access-Control-Allow-Methods": "POST, OPTIONS"
        }
      });
    } catch (e) { lastErr = e; }
  }
  return new Response("星火调用失败: " + (lastErr && lastErr.message), { status: 502, headers: { "Access-Control-Allow-Origin": "*" } });
}

export default {
  async fetch(request) {
    const url = new URL(request.url);
    if (request.method === "OPTIONS") {
      return new Response(null, { status: 204, headers: {
        "Access-Control-Allow-Origin": "*",
        "Access-Control-Allow-Headers": "Content-Type",
        "Access-Control-Allow-Methods": "POST, OPTIONS"
      }});
    }
    if (request.method === "POST" && (url.pathname === "/" || url.pathname === "/chat")) {
      return handlePost(request);
    }
    return new Response("MiniQA Spark Proxy. POST {context,history,question}", { status: 200, headers: { "Access-Control-Allow-Origin": "*" } });
  }
};
