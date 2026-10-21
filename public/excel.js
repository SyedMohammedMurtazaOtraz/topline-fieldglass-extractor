/* Writes one employee into the Field Glass workbook (ExcelJS), keeping the template's formatting.
   Browser: window.FGExcel. Node: module.exports. */
(function (root) {
  const norm = (s) => String(s == null ? "" : s.richText ? s.richText.map((r) => r.text).join("") : s)
    .toLowerCase().replace(/[^a-z0-9]/g, "");

  // Header text (normalised, prefix match) → field key. Falls back to the fixed column letters.
  const HEADERS = [
    ["sno", "sno"], ["maincontract", "main_contract"], ["subcontractor", "subcontractor"],
    ["nameasemiratesidinenglish", "name_en"], ["nameasemiratesidinarabic", "name_ar"], ["email", "email"],
    ["emiratesid", "emirates_id"], ["maritalstatus", "marital_status"], ["nationality", "nationality"], ["dob", "dob"],
    ["bloodgroup", "blood_group"], ["gender", "gender"], ["placeofissueresidence", "residence_place"],
    ["passportno", "passport_no"], ["educationdegree", "education"], ["major", "major"], ["religion", "religion"],
    ["sect", "sect"], ["jobtitle", "job_title"], ["cellphone", "cell_phone"], ["emergencycontact", "emergency_phone"],
    ["h2straining", "h2s_validity"], ["adnoclifesaving", "lifesaving_validity"], ["adnocmedical", "medical_validity"],
    ["hseinduction", "induction_date"], ["certificateno", "certificate_no"], ["workorder", "work_order"],
    ["fieldglassstatus", "fg_status"], ["loaid", "loa_id"], ["location", "location"], ["remarks", "remarks"],
  ];

  function columnMap(ws, FIELDS) {
    for (let r = 1; r <= 6; r++) {
      const row = ws.getRow(r);
      const map = {};
      row.eachCell({ includeEmpty: false }, (cell, c) => {
        const h = norm(cell.value);
        // Longest matching prefix wins (so "emiratesid" doesn't swallow the name columns).
        let best = null;
        HEADERS.forEach(([p, key]) => { if (h.startsWith(p) && (!best || p.length > best[0].length)) best = [p, key]; });
        if (best && !map[best[1]]) map[best[1]] = c;
      });
      if (map.name_en && map.emirates_id) return { headerRow: r, map };
    }
    // Fallback: template layout
    const map = {};
    FIELDS.forEach((f) => { map[f.key] = colNum(f.col); });
    return { headerRow: 2, map };
  }

  function colNum(letters) {
    return letters.split("").reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
  }

  const cellText = (v) => {
    if (v == null) return "";
    if (typeof v === "object" && v.richText) return v.richText.map((r) => r.text).join("");
    if (typeof v === "object" && "result" in v) return String(v.result ?? "");
    return String(v);
  };

  /** Finds the row for this employee: same Emirates ID → update; else first row with no name. */
  function findRow(ws, cm, eid) {
    const { headerRow, map } = cm;
    const last = Math.max(ws.rowCount, headerRow + 1);
    const clean = (eid || "").replace(/\D/g, "");
    if (clean) {
      for (let r = headerRow + 1; r <= last; r++) {
        if (cellText(ws.getCell(r, map.emirates_id).value).replace(/\D/g, "") === clean) return { row: r, existing: true };
      }
    }
    for (let r = headerRow + 1; r <= last + 1; r++) {
      if (!cellText(ws.getCell(r, map.name_en).value).trim()) return { row: r, existing: false };
    }
    return { row: last + 1, existing: false };
  }

  function nextSerial(ws, cm, row) {
    let max = 0;
    for (let r = cm.headerRow + 1; r <= ws.rowCount; r++) {
      if (r === row) continue;
      const n = Number(cellText(ws.getCell(r, cm.map.sno).value));
      if (Number.isFinite(n) && n > max) max = n;
    }
    return max + 1;
  }

  const toDate = (s) => {
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s || "");
    return m ? new Date(Date.UTC(+m[1], +m[2] - 1, +m[3])) : null;
  };

  /** values: merged record (from FGMerge). Returns { row, existing }. */
  function writeEmployee(wb, values, FIELDS) {
    const ws = wb.worksheets[0];
    const cm = columnMap(ws, FIELDS);
    const target = findRow(ws, cm, values.emirates_id);
    const r = target.row;
    // Copy style from the row above if this row has never been formatted.
    const styleFrom = r > cm.headerRow + 1 ? ws.getRow(r - 1) : null;

    FIELDS.forEach((f) => {
      const c = cm.map[f.key];
      if (!c) return;
      const cell = ws.getCell(r, c);
      if (styleFrom && (!cell.style || !Object.keys(cell.style).length)) cell.style = { ...styleFrom.getCell(c).style };

      let v = values[f.key];
      if (f.key === "sno") {
        const existing = cellText(cell.value);
        cell.value = existing ? cell.value : nextSerial(ws, cm, r);
        return;
      }
      if (v === undefined) return;
      if (v === "" || v === null) {
        if (!target.existing) cell.value = null;
        return;
      }
      if (f.type === "date") {
        const d = toDate(v);
        cell.value = d || v;
        if (d && (!cell.numFmt || cell.numFmt === "General" || cell.numFmt === "@")) cell.numFmt = "dd/mm/yyyy";
      } else if (f.key === "emirates_id" && /^\d{15}$/.test(v)) {
        cell.value = Number(v); // the sheet stores IDs as numbers (format "0")
        if (!cell.numFmt || cell.numFmt === "General") cell.numFmt = "0";
      } else {
        cell.value = String(v);
      }
    });
    return target;
  }

  const api = { writeEmployee, columnMap, findRow };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.FGExcel = api;
})(typeof window !== "undefined" ? window : globalThis);
