/**
 * SiPaMeL - Cloudflare Worker Reverse Proxy
 * Target: sipamel.onrender.com
 * Fitur:
 * - Persistent custom link (*.workers.dev)
 * - Auto-rewrite Location header (agar login dan redirect tetap di domain worker)
 * - Caching cerdas untuk aset statis (CSS, JS, gambar, font)
 * - Mendukung upload berkas (POST, multipart/form-data) dan session cookie
 */

const TARGET_HOST = "sipamel.onrender.com";

export default {
  async fetch(request, env, ctx) {
    const url = new URL(request.url);
    const workerHost = url.host;

    // Arahkan request ke server target di Render
    url.hostname = TARGET_HOST;
    url.protocol = "https:";
    url.port = "";

    // Salin header asli dan sesuaikan Host
    const requestHeaders = new Headers(request.headers);
    requestHeaders.set("Host", TARGET_HOST);
    requestHeaders.set("X-Forwarded-Host", workerHost);
    requestHeaders.set("X-Forwarded-Proto", "https");

    // Opsi Caching: Cache aset statis untuk menghemat beban CPU & RAM Render
    const isStaticAsset = /\.(css|js|woff2?|ttf|eot|svg|png|jpg|jpeg|gif|ico|webp)$/i.test(url.pathname);
    const fetchOptions = {
      method: request.method,
      headers: requestHeaders,
      body: ["GET", "HEAD"].includes(request.method) ? undefined : request.body,
      redirect: "manual" // Kita tangani redirect manual agar rewrite Location bekerja mulus
    };

    if (isStaticAsset && request.method === "GET") {
      fetchOptions.cf = {
        cacheEverything: true,
        cacheTtl: 86400 // Cache 24 jam di Edge Cloudflare
      };
    }

    try {
      const response = await fetch(url.toString(), fetchOptions);

      // Salin response headers
      const responseHeaders = new Headers(response.headers);

      // Jika ada redirect (301, 302, 303, 307, 308), ubah alamat Location ke worker host
      const location = responseHeaders.get("Location");
      if (location) {
        const rewrittenLocation = location.replace(`https://${TARGET_HOST}`, `https://${workerHost}`)
                                          .replace(`http://${TARGET_HOST}`, `https://${workerHost}`);
        responseHeaders.set("Location", rewrittenLocation);
      }

      // Pastikan cookie domain sesuai
      const setCookie = responseHeaders.get("Set-Cookie");
      if (setCookie) {
        responseHeaders.set("Set-Cookie", setCookie.replace(new RegExp(`Domain=${TARGET_HOST}`, "gi"), `Domain=${workerHost}`));
      }

      // Tambahkan header identifikasi Cloudflare SiPaMeL
      responseHeaders.set("X-Powered-By", "SiPaMeL-Cloudflare-Worker");

      return new Response(response.body, {
        status: response.status,
        statusText: response.statusText,
        headers: responseHeaders
      });

    } catch (err) {
      return new Response(
        `<html><body style="font-family:sans-serif;text-align:center;padding:50px;">
          <h2>Server Sedang Memulai (Warming Up)...</h2>
          <p>Server Render sedang aktif. Silakan muat ulang halaman dalam 20 detik.</p>
          <p><small>Error: ${err.message}</small></p>
        </body></html>`,
        {
          status: 502,
          headers: { "Content-Type": "text/html; charset=utf-8" }
        }
      );
    }
  }
};
