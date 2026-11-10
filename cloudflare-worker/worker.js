/**
 * SiPaMeL - Cloudflare Worker Smart Reverse Proxy & Auto-Failover
 *
 * Fitur:
 * 1. Auto-Failover Zero-502: Jika Render sedang tidur (cold start) / 502 / 503 / 504,
 *    otomatis dialihkan seketika ke Railway (2 vCPU, 1 GB RAM) sambil membangunkan Render.
 * 2. Smart AI Routing: Endpoint AI OCR berat (/api/ocr/, ocr-scan-uploads) diprioritaskan
 *    ke Railway agar selesai dalam 2 detik tanpa timeout 502 Render.
 * 3. Keep-Alive Cron (*/10 * * * *): Menjaga Render tetap terjaga 24/7 (Zero Cold Start).
 * 4. Smart Edge Caching: Cache aset statis CSS/JS/Gambar selama 24 jam di Edge Jakarta.
 * 5. Friendly Auto-Reload Page: Jika kedua backend sedang restart bersamaan, tampilkan halaman
 *    pemulihan ramah dengan auto-reload otomatis setiap 5 detik (bukan halaman error 502).
 */

const RENDER_HOST = "sipamel.onrender.com";
const RAILWAY_HOST = "sipamel.up.railway.app";

async function fetchBackend(targetHost, originalUrl, method, originalHeaders, bodyBuffer, workerHost, isStaticAsset, timeoutMs) {
  const targetUrl = new URL(originalUrl.toString());
  targetUrl.hostname = targetHost;
  targetUrl.protocol = "https:";
  targetUrl.port = "";

  const requestHeaders = new Headers(originalHeaders);
  requestHeaders.set("Host", targetHost);
  requestHeaders.set("X-Forwarded-Host", workerHost);
  requestHeaders.set("X-Forwarded-Proto", "https");

  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);

  const fetchOptions = {
    method: method,
    headers: requestHeaders,
    body: bodyBuffer !== null ? bodyBuffer : undefined,
    redirect: "manual",
    signal: controller.signal
  };

  if (isStaticAsset && method === "GET") {
    fetchOptions.cf = {
      cacheEverything: true,
      cacheTtl: 86400
    };
  }

  try {
    const response = await fetch(targetUrl.toString(), fetchOptions);
    return response;
  } finally {
    clearTimeout(timer);
  }
}

function buildClientResponse(response, activeHost, workerHost) {
  const responseHeaders = new Headers(response.headers);

  // Rewrite Location redirect URL agar browser tetap di domain Cloudflare Worker
  const location = responseHeaders.get("Location");
  if (location) {
    let rewritten = location;
    for (const h of [RENDER_HOST, RAILWAY_HOST]) {
      rewritten = rewritten
        .replace(`https://${h}`, `https://${workerHost}`)
        .replace(`http://${h}`, `https://${workerHost}`);
    }
    responseHeaders.set("Location", rewritten);
  }

  // Rewrite Cookie Domain jika diset eksplisit oleh origin
  if (typeof response.headers.getSetCookie === "function") {
    const cookies = response.headers.getSetCookie();
    if (cookies && cookies.length > 0) {
      responseHeaders.delete("Set-Cookie");
      for (const c of cookies) {
        responseHeaders.append(
          "Set-Cookie",
          c.replace(new RegExp(`Domain=${activeHost}`, "gi"), `Domain=${workerHost}`)
        );
      }
    }
  } else {
    const setCookie = responseHeaders.get("Set-Cookie");
    if (setCookie) {
      responseHeaders.set(
        "Set-Cookie",
        setCookie.replace(new RegExp(`Domain=${activeHost}`, "gi"), `Domain=${workerHost}`)
      );
    }
  }

  responseHeaders.set("X-Powered-By", "SiPaMeL-Cloudflare-Worker");
  responseHeaders.set("X-Backend-Origin", activeHost);

  return new Response(response.body, {
    status: response.status,
    statusText: response.statusText,
    headers: responseHeaders
  });
}

export default {
  async fetch(request, env, ctx) {
    const url = new URL(request.url);
    const workerHost = url.host;
    const method = request.method;
    const isStaticAsset = /\.(css|js|woff2?|ttf|eot|svg|png|jpg|jpeg|gif|ico|webp)$/i.test(url.pathname);
    const isHeavyOcr = url.pathname.includes("/ocr") || url.pathname.includes("ocr-scan-uploads");
    const isApiCall = url.pathname.startsWith("/api/") || (request.headers.get("Accept") || "").includes("application/json");

    // Buffer body sekali saja untuk POST/PUT/PATCH agar aman jika perlu failover ke backend kedua
    let bodyBuffer = null;
    if (!["GET", "HEAD"].includes(method)) {
      bodyBuffer = await request.arrayBuffer();
    }

    // Tentukan urutan server:
    // - Endpoint OCR berat diprioritaskan ke Railway (2 vCPU, 1 GB RAM, cepat tanpa 502)
    // - Halaman/fitur reguler diprioritaskan ke Render (Gratis permanen), dengan failover instan ke Railway
    const primaryHost = isHeavyOcr ? RAILWAY_HOST : RENDER_HOST;
    const secondaryHost = isHeavyOcr ? RENDER_HOST : RAILWAY_HOST;
    const primaryTimeoutMs = isHeavyOcr ? 55000 : (url.pathname.includes("export") ? 25000 : 12000);
    const secondaryTimeoutMs = isHeavyOcr ? 55000 : 30000;

    // 1. Coba Server Utama (Primary)
    try {
      const resp1 = await fetchBackend(
        primaryHost, url, method, request.headers, bodyBuffer, workerHost, isStaticAsset, primaryTimeoutMs
      );

      // Jika Primary sukses (bukan 502/503/504), langsung kembalikan ke pengguna
      if (![502, 503, 504].includes(resp1.status)) {
        return buildClientResponse(resp1, primaryHost, workerHost);
      }
      // Jika Render sedang 502/503/504, kirim ping wakeup latar belakang
      if (primaryHost === RENDER_HOST) {
        ctx.waitUntil(fetch(`https://${RENDER_HOST}/health`).catch(() => {}));
      }
    } catch (err1) {
      // Primary timeout atau network error: bangunkan Render di background jika Render yang gagal
      if (primaryHost === RENDER_HOST) {
        ctx.waitUntil(fetch(`https://${RENDER_HOST}/health`).catch(() => {}));
      }
    }

    // 2. Auto-Failover ke Server Cadangan (Secondary)
    try {
      const resp2 = await fetchBackend(
        secondaryHost, url, method, request.headers, bodyBuffer, workerHost, isStaticAsset, secondaryTimeoutMs
      );

      if (![502, 503, 504].includes(resp2.status)) {
        return buildClientResponse(resp2, secondaryHost, workerHost);
      }
    } catch (err2) {
      // Secondary juga gagal atau timeout
    }

    // 3. Jika kedua server sementara tidak tersedia (sedang restart / warming up bersamaan)
    if (isApiCall) {
      return new Response(
        JSON.stringify({
          status: "error",
          message: "Server sedang memulai ulang atau memproses antrean. Silakan coba lagi dalam beberapa detik."
        }),
        {
          status: 503,
          headers: {
            "Content-Type": "application/json; charset=utf-8",
            "Retry-After": "5"
          }
        }
      );
    }

    return new Response(
      `<!DOCTYPE html>
      <html lang="id">
      <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <meta http-equiv="refresh" content="5">
        <title>SiPaMeL - Memulai Layanan</title>
        <style>
          body { font-family: system-ui, -apple-system, sans-serif; background: #f8fafc; color: #1e293b; display: flex; align-items: center; justify-content: center; min-height: 100vh; margin: 0; }
          .card { background: white; padding: 2.5rem; border-radius: 16px; box-shadow: 0 10px 25px rgba(0,0,0,0.06); text-align: center; max-width: 420px; width: 90%; }
          .spinner { width: 44px; height: 44px; border: 4px solid #e2e8f0; border-top-color: #4f46e5; border-radius: 50%; animation: spin 1s linear infinite; margin: 0 auto 1.25rem; }
          @keyframes spin { to { transform: rotate(360deg); } }
          h2 { margin: 0 0 0.5rem; font-size: 1.25rem; color: #0f172a; }
          p { margin: 0 0 1.25rem; color: #64748b; font-size: 0.92rem; line-height: 1.5; }
          .btn { display: inline-block; background: #4f46e5; color: white; text-decoration: none; padding: 0.6rem 1.25rem; border-radius: 8px; font-weight: 600; font-size: 0.9rem; }
        </style>
      </head>
      <body>
        <div class="card">
          <div class="spinner"></div>
          <h2>SiPaMeL Sedang Menyiapkan Server</h2>
          <p>Server sedang bangun dari mode hemat daya. Halaman ini akan dimuat ulang secara otomatis dalam 5 detik...</p>
          <a href="${request.url}" class="btn">Muat Ulang Sekarang</a>
        </div>
      </body>
      </html>`,
      {
        status: 503,
        headers: {
          "Content-Type": "text/html; charset=utf-8",
          "Retry-After": "5"
        }
      }
    );
  },

  // Cron Trigger (*/10 * * * *): Menjaga Render agar tidak masuk mode sleep
  async scheduled(event, env, ctx) {
    ctx.waitUntil(
      fetch(`https://${RENDER_HOST}/health`, {
        headers: { "User-Agent": "SiPaMeL-KeepAlive-Cron/1.0" }
      }).catch(() => {})
    );
  }
};
