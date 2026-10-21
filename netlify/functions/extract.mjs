import { verify, json } from "../lib/auth-utils.mjs";

// One call per uploaded file. The browser sends the file already prepared:
// text pages as text, scanned pages / photos as compressed JPEGs.
// Body: { fileName, parts: [{ kind: "text", text, page } | { kind: "image", mediaType, data, page }] }

const DATE = { type: ["string", "null"], description: "YYYY-MM-DD, or null if not shown" };
const STR = { type: ["string", "null"] };

const TOOL = {
  name: "record_document",
  description: "Record the document type and every field visible on this document.",
  input_schema: {
    type: "object",
    properties: {
      doc_type: {
        type: "string",
        enum: [
          "emirates_id",
          "residence_visa",
          "passport",
          "medical_fitness",
          "hse_induction",
          "h2s_certificate",
          "other_certificate",
          "photo",
          "other",
        ],
        description:
          "emirates_id = UAE Resident Identity Card (front/back). residence_visa = UAE 'Residence' permit/sticker (إقامة). passport = passport bio page or other passport pages. medical_fitness = ADNOC/OccuMed medical fitness certificate or medical report. hse_induction = 'ADNOC Onshore HSE Induction & Life Saving Rules' certificate. h2s_certificate = 'H2S Awareness' / SCBA training certificate. photo = a plain passport-style photo with no text.",
      },
      person_name_en: { ...STR, description: "Full name in English exactly as printed (for passports: given names + surname)" },
      person_name_ar: { ...STR, description: "Full name in Arabic script exactly as printed" },
      emirates_id_number: { ...STR, description: "15-digit Emirates ID number, digits only (no dashes)" },
      passport_number: STR,
      date_of_birth: DATE,
      nationality: { ...STR, description: "Country name in English, e.g. India, Pakistan, Nepal" },
      gender: { type: ["string", "null"], enum: ["Male", "Female", null] },
      occupation: { ...STR, description: "Profession / occupation / job title as printed in English" },
      employer: STR,
      place_of_issue: { ...STR, description: "Residence permit place of issue (e.g. Abu Dhabi, Dubai). For passports: passport place of issue." },
      issue_date: { ...DATE, description: "Issue date of the document / certificate (YYYY-MM-DD)" },
      expiry_date: {
        ...DATE,
        description: "Expiry / validity / 'Certificate Validity' / 'Medical expiry date' (YYYY-MM-DD)",
      },
      certificate_number: { ...STR, description: "Certificate No. on training certificates" },
      course_title: STR,
      blood_group: {
        ...STR,
        description: "Blood group with Rh sign, e.g. 'B+', 'AB+', 'O-'. Combine ABO group and Rh type (POSITIVE → +, NEGATIVE → -).",
      },
      medical_result: { ...STR, description: "e.g. 'FIT to work'" },
      marital_status: STR,
      religion: STR,
      notes: {
        type: "string",
        description: "Short note on anything unclear, unreadable, handwritten or contradictory. Empty string if none.",
      },
      confidence: {
        type: "object",
        description: "For each field you filled, 'high' or 'low'. Use 'low' when the text was blurred, cropped or guessed.",
        additionalProperties: { type: "string", enum: ["high", "low"] },
      },
    },
    required: ["doc_type", "notes", "confidence"],
  },
};

const SYSTEM = `You read UAE employee onboarding documents for TopLine Middle East Pipelines Contracting (an ADNOC/CPECC subcontractor) and record their fields.
Rules:
- Copy values exactly as printed. Never invent or infer a value that is not on the document; use null.
- Dates: convert any format (16/05/1985, 25.07.2028, 24th Jul 2028) to YYYY-MM-DD. UAE documents use day/month/year.
- The machine-readable zone (MRZ) can confirm numbers and dates; prefer the printed field when they agree, and mention a disagreement in notes.
- Emirates ID number: 15 digits starting with 784, no dashes.
- For a medical report, the expiry is the "MEDICAL EXPIRY DATE" / "EXPIRY" on the fitness certificate, and blood group comes from the laboratory "BLOOD TYPING; ABO" section.
- For training certificates, issue_date is "Issue Date" and expiry_date is "Certificate Validity".`;

export default async (req) => {
  if (req.method !== "POST") return json({ error: "Method not allowed" }, 405);
  if (!verify(req)) return json({ error: "Session expired — please sign in again." }, 401);
  if (!process.env.ANTHROPIC_API_KEY) return json({ error: "Server is missing ANTHROPIC_API_KEY." }, 500);

  let body;
  try {
    body = await req.json();
  } catch {
    return json({ error: "Invalid request." }, 400);
  }
  const parts = Array.isArray(body.parts) ? body.parts.slice(0, 40) : [];
  if (!parts.length) return json({ error: "No content in file." }, 400);

  const content = [];
  for (const p of parts) {
    if (p.kind === "image" && p.data) {
      content.push({ type: "text", text: `Page ${p.page ?? ""} (image):` });
      content.push({ type: "image", source: { type: "base64", media_type: p.mediaType || "image/jpeg", data: p.data } });
    } else if (p.kind === "text" && p.text) {
      content.push({ type: "text", text: `Page ${p.page ?? ""} (text layer):\n${String(p.text).slice(0, 20000)}` });
    }
  }
  content.push({
    type: "text",
    text: `File name: ${String(body.fileName || "").slice(0, 200)}\nIdentify this document and record its fields with the record_document tool.`,
  });

  const res = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: {
      "content-type": "application/json",
      "x-api-key": process.env.ANTHROPIC_API_KEY,
      "anthropic-version": "2023-06-01",
    },
    body: JSON.stringify({
      model: process.env.ANTHROPIC_MODEL || "claude-sonnet-4-5",
      max_tokens: 1500,
      system: SYSTEM,
      tools: [TOOL],
      tool_choice: { type: "tool", name: TOOL.name },
      messages: [{ role: "user", content }],
    }),
  });

  if (!res.ok) {
    const t = await res.text();
    console.error("Anthropic error", res.status, t.slice(0, 500));
    return json({ error: `AI service error (${res.status}). Try again.` }, 502);
  }
  const data = await res.json();
  const tool = (data.content || []).find((c) => c.type === "tool_use");
  if (!tool) return json({ error: "AI returned no result." }, 502);
  return json({ fileName: body.fileName, result: tool.input });
};

export const config = { path: "/api/extract" };
