const PROXY_SHARED_SECRET = Deno.env.get("PROXY_SHARED_SECRET") ?? "";
const ALLOWED_ORIGIN = Deno.env.get("ALLOWED_ORIGIN") ?? "*";
const DROPBOX_REFRESH_TOKEN = Deno.env.get("DROPBOX_REFRESH_TOKEN") ?? "";
const DROPBOX_APP_KEY = Deno.env.get("DROPBOX_APP_KEY") ?? "";
const DROPBOX_APP_SECRET = Deno.env.get("DROPBOX_APP_SECRET") ?? "";
const DROPBOX_BASE_PATH = "";

async function getAccessToken(): Promise<string> {
  const res = await fetch("https://api.dropbox.com/oauth2/token", {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      grant_type: "refresh_token",
      refresh_token: DROPBOX_REFRESH_TOKEN,
      client_id: DROPBOX_APP_KEY,
      client_secret: DROPBOX_APP_SECRET,
    }),
  });
  const data = await res.json();
  if (!data.access_token) throw new Error("Gagal refresh token: " + JSON.stringify(data));
  return data.access_token;
}

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") {
    return new Response(null, {
      headers: {
        "Access-Control-Allow-Origin": ALLOWED_ORIGIN,
        "Access-Control-Allow-Methods": "POST, OPTIONS",
        "Access-Control-Allow-Headers": "Authorization, Content-Type, x-proxy-secret",
      },
    });
  }
  const secret = req.headers.get("x-proxy-secret") ?? "";
  if (!PROXY_SHARED_SECRET || secret !== PROXY_SHARED_SECRET) {
    return new Response(JSON.stringify({ error: "Unauthorized" }), {
      status: 401,
      headers: { "Content-Type": "application/json" },
    });
  }
  try {
    const accessToken = await getAccessToken();
    const body = await req.json();
    const { action, path, payload, base64 } = body;
    const fileName = path.replace(/^\/+/, "");
    const fullPath = `${DROPBOX_BASE_PATH}/${fileName}`;
    let dropboxUrl = "";
    let dropboxHeaders: Record<string, string> = {
      Authorization: `Bearer ${accessToken}`,
    };
    let dropboxBody: BodyInit | undefined;
    if (action === "restore") {
      dropboxUrl = "https://content.dropboxapi.com/2/files/download";
      dropboxHeaders["Dropbox-API-Arg"] = JSON.stringify({ path: fullPath });
      dropboxHeaders["Content-Type"] = "text/plain";
    } else if (action === "backup") {
      dropboxUrl = "https://content.dropboxapi.com/2/files/upload";
      dropboxHeaders["Dropbox-API-Arg"] = JSON.stringify({ path: fullPath, mode: "overwrite", autorename: false });
      dropboxHeaders["Content-Type"] = "application/octet-stream";
      dropboxBody = JSON.stringify(payload);
    } else if (action === "upload_excel") {
      dropboxUrl = "https://content.dropboxapi.com/2/files/upload";
      dropboxHeaders["Dropbox-API-Arg"] = JSON.stringify({ path: fullPath, mode: "overwrite", autorename: false });
      dropboxHeaders["Content-Type"] = "application/octet-stream";
      const bin = atob(base64);
      const arr = new Uint8Array(bin.length);
      for (let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
      dropboxBody = arr;
    } else {
      return new Response(JSON.stringify({ error: "Invalid action" }), {
        status: 400,
        headers: { "Content-Type": "application/json" },
      });
    }
    const dropboxRes = await fetch(dropboxUrl, {
      method: "POST",
      headers: dropboxHeaders,
      body: dropboxBody,
    });
    const result = await dropboxRes.text();
    return new Response(JSON.stringify({ ok: true, result, debug_path: fullPath }), {
      headers: {
        "Content-Type": "application/json",
        "Access-Control-Allow-Origin": ALLOWED_ORIGIN,
      },
    });
  } catch (err) {
    return new Response(JSON.stringify({ ok: false, msg: String(err) }), {
      status: 500,
      headers: { "Content-Type": "application/json", "Access-Control-Allow-Origin": ALLOWED_ORIGIN },
    });
  }
});
