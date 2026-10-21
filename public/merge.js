/* Merges per-document AI results into one Field Glass employee record + checks.
   Works in the browser (window.FGMerge) and in Node (module.exports) for tests. */
(function (root) {
  const DOC_LABEL = {
    emirates_id: "Emirates ID",
    residence_visa: "Residence Visa",
    passport: "Passport",
    medical_fitness: "ADNOC Medical",
    hse_induction: "HSE Induction / Life Saving",
    h2s_certificate: "H2S Certificate",
    other_certificate: "Other certificate",
    photo: "Photo",
    other: "Other",
  };
  // The documents HR is expected to provide for each employee.
  const EXPECTED = ["emirates_id", "residence_visa", "passport", "medical_fitness", "hse_induction", "h2s_certificate"];

  // Defaults copied from the existing Field Glass sheet.
  const DEFAULTS = {
    main_contract: "CPECC",
    subcontractor: "TOP LINE MIDDLEEAST PIPELINES CONTRACTING L.L.C",
    email: "hr@toplineme.com",
    marital_status: "Married",
    education: "Secondary School",
    major: "12th",
    cell_phone: "00971564338509",
    emergency_phone: "00971564338509",
  };

  // key, label, column letter, type, sources (in priority order: [docType, aiField])
  const FIELDS = [
    { key: "sno", label: "S.No", col: "A", type: "auto" },
    { key: "main_contract", label: "Main Contract", col: "B", type: "text" },
    { key: "subcontractor", label: "Subcontractor", col: "C", type: "text" },
    { key: "name_en", label: "Name (Emirates ID, English)", col: "D", type: "text", required: true, cross: true,
      src: [["emirates_id", "person_name_en"], ["residence_visa", "person_name_en"], ["medical_fitness", "person_name_en"], ["passport", "person_name_en"]] },
    { key: "name_ar", label: "Name (Emirates ID, Arabic)", col: "E", type: "text", rtl: true, required: true,
      src: [["emirates_id", "person_name_ar"], ["residence_visa", "person_name_ar"]] },
    { key: "email", label: "Email", col: "F", type: "text" },
    { key: "emirates_id", label: "Emirates ID", col: "G", type: "eid", required: true, cross: true,
      src: [["emirates_id", "emirates_id_number"], ["residence_visa", "emirates_id_number"], ["medical_fitness", "emirates_id_number"]] },
    { key: "marital_status", label: "Marital Status", col: "H", type: "select", options: ["Married", "Single", "Divorced", "Widowed"], manual: true },
    { key: "nationality", label: "Nationality", col: "I", type: "text", required: true,
      src: [["emirates_id", "nationality"], ["passport", "nationality"], ["medical_fitness", "nationality"], ["residence_visa", "nationality"]] },
    { key: "dob", label: "Date of Birth", col: "J", type: "date", required: true, cross: true,
      src: [["emirates_id", "date_of_birth"], ["passport", "date_of_birth"], ["medical_fitness", "date_of_birth"]] },
    { key: "blood_group", label: "Blood Group", col: "K", type: "select", options: ["A+", "A-", "B+", "B-", "AB+", "AB-", "O+", "O-"], required: true,
      src: [["medical_fitness", "blood_group"], ["other_certificate", "blood_group"]] },
    { key: "gender", label: "Gender", col: "L", type: "select", options: ["Male", "Female"], required: true,
      src: [["emirates_id", "gender"], ["passport", "gender"], ["medical_fitness", "gender"]] },
    { key: "residence_place", label: "Place of issue (Residence Permit)", col: "M", type: "text", required: true,
      src: [["residence_visa", "place_of_issue"]] },
    { key: "passport_no", label: "Passport No", col: "N", type: "text", required: true, cross: true,
      src: [["passport", "passport_number"], ["residence_visa", "passport_number"], ["medical_fitness", "passport_number"]] },
    { key: "education", label: "Education Degree", col: "O", type: "select",
      options: ["Secondary School", "Diploma", "Bachelor", "Master", "Primary School", "None"], manual: true },
    { key: "major", label: "Major", col: "P", type: "text", manual: true },
    { key: "religion", label: "Religion", col: "Q", type: "select", options: ["Muslim", "Hindu", "Christian", "Sikh", "Buddhist", "Other"], manual: true,
      src: [["medical_fitness", "religion"], ["residence_visa", "religion"]] },
    { key: "sect", label: "Sect (Sunni / Shia)", col: "R", type: "select", options: ["", "Sunni", "Sheaa"], manual: true },
    { key: "job_title", label: "Job Title", col: "S", type: "text", required: true,
      src: [["residence_visa", "occupation"], ["emirates_id", "occupation"], ["medical_fitness", "occupation"]] },
    { key: "cell_phone", label: "Cell Phone", col: "T", type: "text" },
    { key: "emergency_phone", label: "Emergency Contact", col: "U", type: "text" },
    { key: "h2s_validity", label: "H2S Training Validity", col: "V", type: "date", required: true, expiry: true,
      src: [["h2s_certificate", "expiry_date"]] },
    { key: "lifesaving_validity", label: "ADNOC Life Saving Cert. Validity", col: "W", type: "date", required: true, expiry: true,
      src: [["hse_induction", "expiry_date"]] },
    { key: "medical_validity", label: "ADNOC Medical Validity", col: "X", type: "date", required: true, expiry: true,
      src: [["medical_fitness", "expiry_date"]] },
    { key: "induction_date", label: "HSE Induction Date", col: "Y", type: "date", required: true,
      src: [["hse_induction", "issue_date"]] },
    { key: "certificate_no", label: "Certificate No. (HSE Induction)", col: "Z", type: "text", required: true,
      src: [["hse_induction", "certificate_number"]] },
    { key: "work_order", label: "Work Order ID / Fieldglass", col: "AA", type: "text", manual: true },
    { key: "fg_status", label: "Fieldglass Status", col: "AB", type: "text", manual: true },
    { key: "loa_id", label: "LOA ID No", col: "AC", type: "text", manual: true },
    { key: "location", label: "Location", col: "AD", type: "text", manual: true },
    { key: "remarks", label: "Remarks", col: "AE", type: "text", manual: true },
  ];

  const NATION = {
    INDIAN: "India", INDIA: "India", "REPUBLIC OF INDIA": "India",
    PAKISTANI: "Pakistan", PAKISTAN: "Pakistan", NEPALI: "Nepal", NEPALESE: "Nepal", NEPAL: "Nepal",
    BANGLADESHI: "Bangladesh", BANGLADESH: "Bangladesh", "SRI LANKAN": "Sri Lanka", "SRI LANKA": "Sri Lanka",
    FILIPINO: "Philippines", PHILIPPINES: "Philippines", EGYPTIAN: "Egypt", EGYPT: "Egypt",
    SUDANESE: "Sudan", SUDAN: "Sudan", SYRIAN: "Syria", SYRIA: "Syria", JORDANIAN: "Jordan", JORDAN: "Jordan",
    UGANDAN: "Uganda", UGANDA: "Uganda", KENYAN: "Kenya", KENYA: "Kenya", GHANAIAN: "Ghana", GHANA: "Ghana",
    CHINESE: "China", CHINA: "China", AFGHAN: "Afghanistan", AFGHANISTAN: "Afghanistan", YEMENI: "Yemen", YEMEN: "Yemen",
  };

  const titleCase = (s) =>
    String(s).toLowerCase().replace(/\s+/g, " ").trim().replace(/(^|[\s\-/(])([a-z])/g, (m, a, b) => a + b.toUpperCase());

  function normalise(key, v) {
    if (v === null || v === undefined) return null;
    let s = String(v).trim();
    if (!s || /^(null|n\/?a|none|-)$/i.test(s)) return null;
    switch (key) {
      case "emirates_id": {
        const d = s.replace(/\D/g, "");
        return d || null;
      }
      case "passport_no":
        return s.replace(/\s+/g, "").toUpperCase();
      case "name_en":
        return titleCase(s.replace(/^(mr|mrs|ms)\.?\s+/i, ""));
      case "nationality":
        return NATION[s.toUpperCase()] || titleCase(s);
      case "job_title":
        return titleCase(s);
      case "residence_place":
        return titleCase(s.replace(/[؀-ۿ]+/g, "").trim()) || s;
      case "gender":
        return /^f/i.test(s) ? "Female" : /^m/i.test(s) ? "Male" : s;
      case "blood_group": {
        const t = s.toUpperCase().replace(/["'\s]/g, "").replace("POSITIVE", "+").replace("NEGATIVE", "-").replace("VE", "");
        const m = t.match(/^(AB|A|B|O)\s*([+-])?/);
        return m ? m[1] + (m[2] || "") : s;
      }
      case "religion":
        return titleCase(s);
      default:
        return s;
    }
  }

  const comparable = (key, v) => {
    if (v == null) return null;
    if (key === "name_en") return v.toUpperCase().replace(/[^A-Z]/g, "");
    return String(v).toUpperCase();
  };

  function daysBetween(a, b) {
    return Math.round((Date.parse(a) - Date.parse(b)) / 86400000);
  }

  /** docs: [{ fileName, result }] → { values, meta, docs, issues } */
  function merge(docs, today) {
    today = today || new Date().toISOString().slice(0, 10);
    const byType = {};
    docs.forEach((d) => {
      const t = (d.result && d.result.doc_type) || "other";
      (byType[t] = byType[t] || []).push(d);
    });

    const values = {};
    const meta = {}; // key -> { source, file, confidence, conflict: [...], status }
    const issues = [];

    FIELDS.forEach((f) => {
      meta[f.key] = { status: "ok" };
      if (f.type === "auto") return;
      const found = [];
      (f.src || []).forEach(([type, field]) => {
        (byType[type] || []).forEach((d) => {
          const raw = d.result[field];
          const v = normalise(f.key, raw);
          if (v) found.push({ v, type, file: d.fileName, conf: (d.result.confidence || {})[field] || "high" });
        });
      });
      if (found.length) {
        const best = found[0];
        values[f.key] = best.v;
        Object.assign(meta[f.key], { source: DOC_LABEL[best.type], file: best.file, confidence: best.conf });
        if (best.conf === "low") meta[f.key].status = "warn";
        if (f.cross) {
          const others = found.filter((x) => comparable(f.key, x.v) !== comparable(f.key, best.v));
          // Names: passports/medical often drop the father's name, so only flag when one isn't contained in the other.
          const real = f.key === "name_en"
            ? others.filter((x) => { const a = comparable(f.key, x.v), b = comparable(f.key, best.v); return !a.includes(b) && !b.includes(a); })
            : others;
          if (real.length) {
            meta[f.key].status = "warn";
            meta[f.key].conflict = real.map((x) => `${DOC_LABEL[x.type]}: ${x.v}`);
            issues.push({ level: "warn", kind: "conflict", field: f.key, text: `${f.label} differs between documents (${[best].concat(real).map((x) => `${DOC_LABEL[x.type]}: ${x.v}`).join(" · ")})` });
          }
        }
      } else if (DEFAULTS[f.key] !== undefined) {
        values[f.key] = DEFAULTS[f.key];
        Object.assign(meta[f.key], { source: "Default" });
        if (f.manual) meta[f.key].status = "manual";
      } else {
        values[f.key] = "";
        if (f.required) {
          meta[f.key].status = "missing";
          issues.push({ level: "error", field: f.key, text: `${f.label} not found in the documents` });
        } else if (f.manual) {
          meta[f.key].status = "manual";
        }
      }
    });

    // Sect only applies to Muslims; leave blank otherwise.
    if (values.religion === "Muslim" && !values.sect) { values.sect = "Sunni"; meta.sect = { status: "manual", source: "Default" }; }

    const present = Object.keys(byType);
    const missingDocs = EXPECTED.filter((t) => !present.includes(t)).map((t) => DOC_LABEL[t]);
    missingDocs.forEach((m) => issues.unshift({ level: "error", text: `No ${m} document was found in the upload` }));

    return { values, meta, issues, byType, missingDocs };
  }

  /** Format / age / expiry checks on the (possibly edited) values. Mutates meta + issues. */
  function checks(values, meta, issues, today) {
    today = today || new Date().toISOString().slice(0, 10);
    if (values.emirates_id && !/^784\d{12}$/.test(values.emirates_id)) {
      meta.emirates_id.status = "warn";
      issues.push({ level: "warn", field: "emirates_id", text: "Emirates ID should be 15 digits starting with 784" });
    }
    if (values.dob) {
      const age = daysBetween(today, values.dob) / 365.25;
      if (!(age >= 18 && age <= 70)) {
        meta.dob.status = "error";
        issues.push({ level: "error", field: "dob", text: `Date of birth ${values.dob} gives an age of ${Math.floor(age)} — check it` });
      }
    }
    FIELDS.filter((f) => f.expiry).forEach((f) => {
      const v = values[f.key];
      if (!v) return;
      const d = daysBetween(v, today);
      if (d < 0) {
        meta[f.key].status = "error";
        issues.push({ level: "error", field: f.key, text: `${f.label} expired on ${v}` });
      } else if (d <= 30) {
        meta[f.key].status = "warn";
        issues.push({ level: "warn", field: f.key, text: d === 0 ? `${f.label} expires today (${v})` : `${f.label} expires in ${d} day${d === 1 ? "" : "s"} (${v})` });
      }
    });
    FIELDS.filter((f) => f.required).forEach((f) => {
      if (!values[f.key] && !issues.some((i) => i.field === f.key)) {
        meta[f.key].status = "missing";
        issues.push({ level: "error", field: f.key, text: `${f.label} not found — please enter it` });
      }
    });
  }

  /** Applies HR's edits on top of a merge result and re-runs all checks. */
  function evaluate(merged, edits, today) {
    edits = edits || {};
    const values = Object.assign({}, merged.values, edits);
    const meta = {};
    Object.keys(merged.meta).forEach((k) => { meta[k] = Object.assign({}, merged.meta[k]); });
    Object.keys(edits).forEach((k) => {
      if (!meta[k]) meta[k] = {};
      if (edits[k] !== merged.values[k]) { meta[k].status = "edited"; meta[k].edited = true; delete meta[k].conflict; }
    });
    Object.keys(meta).forEach((k) => { if (meta[k].status === "missing" && values[k]) meta[k].status = "ok"; });
    const issues = merged.issues.filter((i) => !i.field || (i.kind === "conflict" && !(meta[i.field] && meta[i.field].edited)));
    checks(values, meta, issues, today);
    const errors = issues.filter((i) => i.level === "error").length;
    const warns = issues.filter((i) => i.level === "warn").length;
    return { values, meta, issues, errors, warns };
  }

  const api = { FIELDS, DEFAULTS, DOC_LABEL, EXPECTED, merge, evaluate, checks, normalise };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.FGMerge = api;
})(typeof window !== "undefined" ? window : globalThis);
