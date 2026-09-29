// Cloudflare Pages Functions —— 把 /api/* 同源代理到 WorkBuddy 线上后端
// 作用：前端部署在 Cloudflare 后，请求同源的 /api/data 由本函数转发到 WorkBuddy，
//       既解决跨域，又保留 WorkBuddy 云端的数据持久化（升级/换设备/换浏览器都不丢数据）。
// 若 WorkBuddy 链接变更，只需修改下方 TARGET_ORIGIN 一行，重新部署即可。
const TARGET_ORIGIN = 'https://d193ce80a56d49fc9594a3cb2c0f0d29.app.workbuddy.link';

export async function onRequest(context) {
  const request = context.request;
  const url = new URL(request.url);
  const targetUrl = TARGET_ORIGIN + url.pathname + url.search;

  // 清理不应转发到后端的 hop-by-hop / 代理头
  const headers = new Headers(request.headers);
  ['host', 'content-length', 'cf-connecting-ip', 'cf-visitor', 'cf-ray',
   'x-forwarded-for', 'x-forwarded-proto', 'x-real-ip'].forEach((h) => headers.delete(h));

  const method = request.method;
  const hasBody = method !== 'GET' && method !== 'HEAD';
  const body = hasBody ? await request.arrayBuffer() : undefined;

  // WorkBuddy 网关偶发 502，做最多 3 次重试
  let lastErr;
  for (let i = 0; i < 3; i++) {
    try {
      const resp = await fetch(targetUrl, { method, headers, body, redirect: 'follow' });
      const out = new Headers(resp.headers);
      out.delete('transfer-encoding');
      out.delete('connection');
      out.delete('keep-alive');
      return new Response(resp.body, { status: resp.status, headers: out });
    } catch (e) {
      lastErr = e;
    }
  }
  return new Response(
    '代理后端失败（可能为 WorkBuddy 网关抖动，请稍后重试）：' +
      (lastErr && lastErr.message ? lastErr.message : String(lastErr)),
    { status: 502 }
  );
}
