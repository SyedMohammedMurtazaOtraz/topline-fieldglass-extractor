import crypto from "node:crypto";
import { sign, json } from "../lib/auth-utils.mjs";

// Users: APP_USERS="hr@toplineme.com:pass1,admin@otraz.ai:pass2"  (or a single shared APP_PASSWORD)
function check(email, password) {
  const users = (process.env.APP_USERS || "")
    .split(",")
    .map((s) => s.trim())
    .filter(Boolean)
    .map((s) => {
      const i = s.indexOf(":");
      return [s.slice(0, i).toLowerCase(), s.slice(i + 1)];
    });
  const eq = (a, b) => a.length === b.length && crypto.timingSafeEqual(Buffer.from(a), Buffer.from(b));
  if (users.length) return users.some(([u, p]) => u === email.toLowerCase() && eq(p, password));
  const shared = process.env.APP_PASSWORD || "";
  return !!shared && eq(shared, password);
}

export default async (req) => {
  if (req.method !== "POST") return json({ error: "Method not allowed" }, 405);
  let body = {};
  try { body = await req.json(); } catch {}
  const email = String(body.email || "").trim();
  const password = String(body.password || "");
  if (!email || !password) return json({ error: "Enter your email and password." }, 400);
  if (!check(email, password)) return json({ error: "Incorrect email or password." }, 401);
  return json({ token: sign(email), user: email });
};

export const config = { path: "/api/auth" };
