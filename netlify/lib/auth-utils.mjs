import crypto from "node:crypto";

// Token signing key: AUTH_SECRET if set, otherwise derived from the login settings + API key.
const secret = () =>
  process.env.AUTH_SECRET ||
  [process.env.APP_USERS, process.env.APP_PASSWORD, process.env.ANTHROPIC_API_KEY].filter(Boolean).join("|");

export function sign(user, hours = 12) {
  const exp = Date.now() + hours * 3600 * 1000;
  const payload = Buffer.from(JSON.stringify({ u: user, exp })).toString("base64url");
  const sig = crypto.createHmac("sha256", secret()).update(payload).digest("base64url");
  return `${payload}.${sig}`;
}

export function verify(req) {
  const h = req.headers.get("authorization") || "";
  const token = h.startsWith("Bearer ") ? h.slice(7) : "";
  const [payload, sig] = token.split(".");
  if (!payload || !sig || !secret()) return null;
  const good = crypto.createHmac("sha256", secret()).update(payload).digest("base64url");
  if (good.length !== sig.length || !crypto.timingSafeEqual(Buffer.from(good), Buffer.from(sig))) return null;
  try {
    const data = JSON.parse(Buffer.from(payload, "base64url").toString());
    return data.exp > Date.now() ? data : null;
  } catch {
    return null;
  }
}

export const json = (body, status = 200) =>
  new Response(JSON.stringify(body), { status, headers: { "content-type": "application/json" } });
