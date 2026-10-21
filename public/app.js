/* Fieldglass Extractor — TopLine. Browser app. */
(() => {
  const { FIELDS, DOC_LABEL, merge, evaluate } = window.FGMerge;
  const $ = (id) => document.getElementById(id);
  const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  pdfjsLib.GlobalWorkerOptions.workerSrc = "vendor/pdf.worker.min.js";

  const ACCEPT = /\.(pdf|jpe?g|png|webp)$/i;
  const MAX_PARALLEL = 3;
  const MAX_PAYLOAD = 4_300_000; // chars of JSON; Netlify request limit is ~6 MB

  const state = {
    token: null,
    user: "",
    wb: null,
    wbName: "Fieldglass template",
    wbDirty: false,
    employees: [],
    active: null,
  };

  /* ───────── storage (safe) ───────── */
  const store = {
    get: (k) => { try { return sessionStorage.getItem(k); } catch { return null; } },
    set: (k, v) => { try { sessionStorage.setItem(k, v); } catch {} },
    del: (k) => { try { sessionStorage.removeItem(k); } catch {} },
  };

  /* ───────── toast ───────── */
  let toastTimer;
  function toast(msg, isErr) {
    const t = $("toast");
    t.textContent = msg;
    t.className = "toast" + (isErr ? " err" : "");
    t.hidden = false;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => (t.hidden = true), 4200);
  }

  /* ───────── auth ───────── */
  function showLogin() {
    $("app").hidden = true;
    $("login").hidden = false;
    $("email").focus();
  }
  async function showApp() {
    $("login").hidden = true;
    $("app").hidden = false;
    $("who").textContent = state.user;
    if (!state.wb) await loadTemplate();
    render();
  }
  $("loginForm").addEventListener("submit", async (e) => {
    e.preventDefault();
    $("loginError").textContent = "";
    $("loginBtn").disabled = true;
    $("loginBtn").textContent = "Signing in…";
    try {
      const r = await fetch("/api/auth", {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify({ email: $("email").value, password: $("password").value }),
      });
      const d = await r.json().catch(() => ({}));
      if (!r.ok) throw new Error(d.error || "Sign-in failed.");
      state.token = d.token;
      state.user = d.user;
      store.set("fg_token", d.token);
      store.set("fg_user", d.user);
      $("password").value = "";
      state.paused = false;
      await showApp();
      pump();
    } catch (err) {
      $("loginError").textContent = err.message;
    } finally {
      $("loginBtn").disabled = false;
      $("loginBtn").textContent = "Sign in";
    }
  });
  $("logoutBtn").addEventListener("click", () => {
    if (state.wbDirty && !confirmLeave()) return;
    store.del("fg_token");
    state.token = null;
    showLogin();
  });
  function confirmLeave() {
    return window.confirm("You have rows that haven't been downloaded yet. Sign out anyway?");
  }
  window.addEventListener("beforeunload", (e) => {
    if (state.wbDirty) { e.preventDefault(); e.returnValue = ""; }
  });

  /* ───────── workbook ───────── */
  async function loadTemplate() {
    const r = await fetch("template.xlsx");
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(await r.arrayBuffer());
    state.wb = wb;
    state.wbName = "Fieldglass template (blank)";
    state.wbDirty = false;
    updateWbPill();
  }
  function wbRowCount() {
    if (!state.wb) return 0;
    const ws = state.wb.worksheets[0];
    const cm = FGExcel.columnMap(ws, FIELDS);
    let n = 0;
    for (let r = cm.headerRow + 1; r <= ws.rowCount; r++) {
      const v = ws.getCell(r, cm.map.name_en).value;
      if (v && String(v.richText ? v.richText.map((x) => x.text).join("") : v).trim()) n++;
    }
    return n;
  }
  function updateWbPill() {
    $("wbName").textContent = state.wbName;
    const n = wbRowCount();
    $("wbCount").textContent = `· ${n} employee${n === 1 ? "" : "s"}`;
    $("downloadWbBtn").disabled = !state.wb;
  }
  $("loadMasterBtn").addEventListener("click", () => $("masterInput").click());
  $("masterInput").addEventListener("change", async (e) => {
    const f = e.target.files[0];
    e.target.value = "";
    if (!f) return;
    try {
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(await f.arrayBuffer());
      const cm = FGExcel.columnMap(wb.worksheets[0], FIELDS);
      if (!cm.map.name_en || !cm.map.emirates_id) throw new Error("That file doesn't look like the Fieldglass sheet.");
      state.wb = wb;
      state.wbName = f.name;
      state.wbDirty = false;
      state.employees.forEach((emp) => { emp.row = null; });
      updateWbPill();
      render();
      toast(`Loaded ${f.name} — new employees will be added to it; existing Emirates IDs are updated.`);
    } catch (err) {
      toast(err.message || "Couldn't read that Excel file.", true);
    }
  });
  async function downloadWorkbook(wb, name) {
    const buf = await wb.xlsx.writeBuffer();
    const blob = new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = name;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 4000);
  }
  const stamp = () => new Date().toISOString().slice(0, 10);
  $("downloadWbBtn").addEventListener("click", async () => {
    const base = state.wbName.startsWith("Fieldglass template") ? "Fieldglass" : state.wbName.replace(/\.xlsx$/i, "");
    await downloadWorkbook(state.wb, `${base} - ${stamp()}.xlsx`);
    state.wbDirty = false;
  });

  /* ───────── intake: files / folders / zips → employees ───────── */
  const dz = $("dropzone");
  $("pickFilesBtn").addEventListener("click", () => $("fileInput").click());
  $("pickFolderBtn").addEventListener("click", () => $("folderInput").click());
  $("fileInput").addEventListener("change", (e) => { intake([...e.target.files].map((f) => ({ file: f, path: f.name }))); e.target.value = ""; });
  $("folderInput").addEventListener("change", (e) => { intake([...e.target.files].map((f) => ({ file: f, path: f.webkitRelativePath || f.name }))); e.target.value = ""; });
  ["dragenter", "dragover"].forEach((ev) => dz.addEventListener(ev, (e) => { e.preventDefault(); dz.classList.add("over"); }));
  ["dragleave", "drop"].forEach((ev) => dz.addEventListener(ev, (e) => { e.preventDefault(); dz.classList.remove("over"); }));
  dz.addEventListener("drop", async (e) => {
    const items = [...(e.dataTransfer.items || [])];
    const entries = items.map((i) => (i.webkitGetAsEntry ? i.webkitGetAsEntry() : null)).filter(Boolean);
    if (entries.length) {
      const out = [];
      await Promise.all(entries.map((en) => walkEntry(en, "", out)));
      intake(out);
    } else {
      intake([...e.dataTransfer.files].map((f) => ({ file: f, path: f.name })));
    }
  });
  function walkEntry(entry, prefix, out) {
    return new Promise((resolve) => {
      if (entry.isFile) {
        entry.file((f) => { out.push({ file: f, path: prefix + f.name }); resolve(); }, () => resolve());
      } else if (entry.isDirectory) {
        const reader = entry.createReader();
        const all = [];
        const read = () => reader.readEntries(async (batch) => {
          if (!batch.length) {
            await Promise.all(all.map((c) => walkEntry(c, prefix + entry.name + "/", out)));
            resolve();
          } else { all.push(...batch); read(); }
        }, () => resolve());
        read();
      } else resolve();
    });
  }

  async function expandZips(items) {
    const out = [];
    for (const it of items) {
      if (/\.zip$/i.test(it.path)) {
        try {
          const zip = await JSZip.loadAsync(it.file);
          const zipBase = it.path.replace(/\.zip$/i, "");
          const files = Object.values(zip.files).filter((z) => !z.dir && !/(^|\/)(__MACOSX|\.)/.test(z.name));
          for (const z of files) {
            const blob = await z.async("blob");
            const name = z.name.split("/").pop();
            out.push({ file: new File([blob], name, { type: guessType(name) }), path: `${zipBase}/${z.name}` });
          }
        } catch {
          toast(`Couldn't open ${it.path}`, true);
        }
      } else out.push(it);
    }
    return out;
  }
  const guessType = (n) => (/\.pdf$/i.test(n) ? "application/pdf" : /\.png$/i.test(n) ? "image/png" : /\.webp$/i.test(n) ? "image/webp" : "image/jpeg");

  /** Groups files by employee folder. Strips shared parent folders first. */
  function groupByEmployee(items) {
    let parts = items.map((it) => ({ it, seg: it.path.split("/").filter(Boolean) }));
    let label = "";
    while (parts.length && parts.every((p) => p.seg.length > 1) && new Set(parts.map((p) => p.seg[0])).size === 1) {
      label = parts[0].seg[0];
      parts.forEach((p) => p.seg.shift());
    }
    const groups = new Map();
    const nested = parts.some((p) => p.seg.length > 1);
    parts.forEach((p) => {
      const key = nested && p.seg.length > 1 ? p.seg[0] : label || "";
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key).push(p.it);
    });
    return [...groups.entries()];
  }
  const prettyFolder = (s) => s.replace(/^\s*\d+\s*[.)\-_]\s*/, "").replace(/[_]+/g, " ").trim();

  async function intake(raw) {
    if (!raw.length) return;
    let items = await expandZips(raw);
    const skipped = items.filter((i) => !ACCEPT.test(i.file.name));
    items = items.filter((i) => ACCEPT.test(i.file.name));
    if (skipped.length) toast(`Skipped ${skipped.length} unsupported file${skipped.length > 1 ? "s" : ""} (${skipped.slice(0, 3).map((s) => s.file.name).join(", ")})`);
    if (!items.length) return;

    const groups = groupByEmployee(items);
    let first = null;
    groups.forEach(([folder, files]) => {
      const emp = newEmployee(folder);
      files.forEach((f) => emp.files.push(newDoc(f.file)));
      state.employees.push(emp);
      if (!first) first = emp;
    });
    state.active = first.id;
    if (groups.length > 1) toast(`Queued ${groups.length} employees from the upload.`);
    render();
    pump();
  }

  let seq = 0;
  function newEmployee(folder) {
    return { id: ++seq, folder: prettyFolder(folder), files: [], merged: null, edits: {}, row: null, added: false };
  }
  function newDoc(file) {
    return { id: ++seq, file, name: file.name, status: "queued", result: null, error: "", typeOverride: null };
  }

  /* ───────── extraction queue ───────── */
  let running = 0;
  function pump() {
    if (state.paused) return;
    while (running < MAX_PARALLEL) {
      const next = nextQueued();
      if (!next) return;
      const [emp, doc] = next;
      doc.status = "working";
      running++;
      renderIfActive(emp);
      runDoc(emp, doc).finally(() => {
        if (doc.status === "queued") { running--; return; }
        running--;
        afterDoc(emp);
        pump();
      });
    }
  }
  function nextQueued() {
    // Active employee first, then the rest in order.
    const order = [...state.employees].sort((a, b) => (a.id === state.active ? -1 : b.id === state.active ? 1 : 0));
    for (const emp of order) for (const doc of emp.files) if (doc.status === "queued") return [emp, doc];
    return null;
  }
  async function runDoc(emp, doc) {
    try {
      const parts = await prepare(doc.file);
      if (!parts.length) throw new Error("No readable content");
      let attempt = 0;
      for (;;) {
        attempt++;
        const r = await fetch("/api/extract", {
          method: "POST",
          headers: { "content-type": "application/json", authorization: `Bearer ${state.token}` },
          body: JSON.stringify({ fileName: doc.name, parts }),
        }).catch(() => null);
        if (r && r.status === 401) {
          // Pause the queue; it resumes after signing in again.
          doc.status = "queued";
          state.paused = true;
          store.del("fg_token");
          showLogin();
          $("loginError").textContent = "Your session expired — sign in to continue reading documents.";
          return;
        }
        const d = r ? await r.json().catch(() => ({})) : {};
        if (r && r.ok && d.result) { doc.result = d.result; doc.status = "done"; return; }
        if (attempt >= 2) throw new Error(d.error || (r ? `Error ${r.status}` : "Network error"));
        await new Promise((res) => setTimeout(res, 1500));
      }
    } catch (err) {
      doc.status = "error";
      doc.error = err.message || "Failed";
    }
  }
  function afterDoc(emp) {
    if (emp.files.every((d) => d.status === "done" || d.status === "error")) remerge(emp);
    renderSidebar();
    renderIfActive(emp);
  }
  function remerge(emp) {
    const docs = emp.files
      .filter((d) => d.status === "done")
      .map((d) => ({ fileName: d.name, result: d.typeOverride ? { ...d.result, doc_type: d.typeOverride } : d.result }));
    emp.merged = merge(docs);
  }

  /* ───────── file preparation (in the browser) ───────── */
  async function prepare(file) {
    const isPdf = /\.pdf$/i.test(file.name) || file.type === "application/pdf";
    let attempts = [[1800, 0.82], [1400, 0.7], [1100, 0.6], [900, 0.5]];
    let canvases = [];
    let textParts = [];
    if (isPdf) {
      const pdf = await pdfjsLib.getDocument({ data: new Uint8Array(await file.arrayBuffer()) }).promise;
      const n = Math.min(pdf.numPages, 30);
      for (let i = 1; i <= n; i++) {
        const page = await pdf.getPage(i);
        const tc = await page.getTextContent();
        const text = tc.items.map((x) => x.str).join(" ").replace(/\s+/g, " ").trim();
        const hasText = text.replace(/\s/g, "").length >= 200;
        if (hasText) textParts.push({ kind: "text", page: i, text });
        // Short documents (ID cards, certificates) always go as images too — layout and stamps matter.
        if ((!hasText || pdf.numPages <= 3) && canvases.length < 8) {
          const vp1 = page.getViewport({ scale: 1 });
          const scale = Math.min(3, 1800 / Math.max(vp1.width, vp1.height));
          const vp = page.getViewport({ scale });
          const c = document.createElement("canvas");
          c.width = Math.round(vp.width);
          c.height = Math.round(vp.height);
          const ctx = c.getContext("2d");
          ctx.fillStyle = "#fff";
          ctx.fillRect(0, 0, c.width, c.height);
          await page.render({ canvasContext: ctx, viewport: vp }).promise;
          canvases.push({ page: i, canvas: c });
        }
      }
      pdf.destroy();
    } else {
      const bmp = await createImageBitmap(file);
      const c = document.createElement("canvas");
      c.width = bmp.width;
      c.height = bmp.height;
      c.getContext("2d").drawImage(bmp, 0, 0);
      canvases.push({ page: 1, canvas: c });
    }
    for (const [side, q] of attempts) {
      const imgs = canvases.map(({ page, canvas }) => ({ kind: "image", page, mediaType: "image/jpeg", data: encode(canvas, side, q) }));
      const parts = [...textParts, ...imgs].sort((a, b) => a.page - b.page);
      if (JSON.stringify(parts).length < MAX_PAYLOAD) return parts;
    }
    throw new Error("File too large to send");
  }
  function encode(src, maxSide, q) {
    const s = Math.min(1, maxSide / Math.max(src.width, src.height));
    const c = document.createElement("canvas");
    c.width = Math.round(src.width * s);
    c.height = Math.round(src.height * s);
    const ctx = c.getContext("2d");
    ctx.fillStyle = "#fff";
    ctx.fillRect(0, 0, c.width, c.height);
    ctx.drawImage(src, 0, 0, c.width, c.height);
    return c.toDataURL("image/jpeg", q).split(",")[1];
  }

  /* ───────── rendering ───────── */
  const activeEmp = () => state.employees.find((e) => e.id === state.active) || null;
  function render() {
    renderSidebar();
    const emp = activeEmp();
    $("uploadView").hidden = !!emp;
    $("empView").hidden = !emp;
    if (emp) renderEmployee(emp);
    updateWbPill();
  }
  function renderIfActive(emp) { if (emp.id === state.active) renderEmployee(emp); }

  function empName(emp) {
    return (emp.merged && emp.merged.values.name_en) || emp.folder || `Employee ${emp.id}`;
  }
  function empState(emp) {
    const total = emp.files.length;
    const done = emp.files.filter((d) => d.status === "done" || d.status === "error").length;
    if (done < total) return { text: `Reading ${done}/${total}`, cls: "spin" };
    if (emp.added) return { text: `Added · row ${emp.row}`, cls: "good" };
    const ev = emp.merged ? evaluate(emp.merged, emp.edits) : null;
    if (!ev) return { text: "—", cls: "" };
    if (ev.errors) return { text: `${ev.errors} to fix`, cls: "error" };
    if (ev.warns) return { text: `${ev.warns} to check`, cls: "warn" };
    return { text: "Ready", cls: "good" };
  }
  function renderSidebar() {
    const ul = $("empList");
    if (!state.employees.length) {
      ul.innerHTML = `<li class="muted small" style="cursor:default">No employees yet.</li>`;
      return;
    }
    ul.innerHTML = state.employees.map((emp) => {
      const s = empState(emp);
      return `<li data-id="${emp.id}" class="${emp.id === state.active ? "active" : ""}">
        <span class="name">${esc(empName(emp))}</span>
        <span class="state">${s.cls === "spin" ? '<span class="spinner"></span>' : `<span class="badge ${s.cls}">${esc(s.text)}</span>`}${s.cls === "spin" ? esc(s.text) : ""}</span>
      </li>`;
    }).join("");
  }
  $("empList").addEventListener("click", (e) => {
    const li = e.target.closest("li[data-id]");
    if (!li) return;
    state.active = Number(li.dataset.id);
    render();
  });
  $("newEmpBtn").addEventListener("click", () => { state.active = null; render(); });

  const DOC_TYPES = Object.keys(DOC_LABEL);
  function renderEmployee(emp) {
    $("empTitle").textContent = empName(emp);
    $("empFolder").textContent = emp.folder && emp.folder !== empName(emp) ? `Folder: ${emp.folder}` : "Employee";

    // Documents
    const done = emp.files.filter((d) => d.status === "done").length;
    $("docProgress").textContent = `${done}/${emp.files.length} read`;
    $("docList").innerHTML = emp.files.map((d) => {
      let right = "";
      if (d.status === "queued") right = `<span class="muted small">Waiting</span>`;
      else if (d.status === "working") right = `<span class="small"><span class="spinner"></span> Reading</span>`;
      else if (d.status === "error") right = `<button class="btn sm" data-retry="${d.id}">Retry</button>`;
      else {
        const t = d.typeOverride || d.result.doc_type;
        right = `<select data-doc="${d.id}" aria-label="Document type">${DOC_TYPES.map((k) => `<option value="${k}" ${k === t ? "selected" : ""}>${esc(DOC_LABEL[k])}</option>`).join("")}</select>`;
      }
      const sub = d.status === "error" ? `<span class="doc-sub" style="color:var(--err)">${esc(d.error)}</span>`
        : d.result && d.result.notes ? `<span class="doc-sub" title="${esc(d.result.notes)}">${esc(d.result.notes)}</span>` : "";
      return `<li><div style="min-width:0"><div class="doc-name" title="${esc(d.name)}">${esc(d.name)}</div>${sub}</div>${right}</li>`;
    }).join("");

    const ready = emp.merged && emp.files.every((d) => d.status === "done" || d.status === "error");
    $("addRowBtn").disabled = !ready;
    $("singleXlsxBtn").disabled = !ready;
    $("addRowBtn").textContent = emp.added ? "Update in workbook" : "Add to workbook";
    $("rowNote").hidden = !emp.added;
    if (emp.added) $("rowNote").textContent = `Written to row ${emp.row} of ${state.wbName}. Remember to download the workbook.`;

    if (!ready) {
      $("issueList").innerHTML = `<li class="muted small">Checks run once every document has been read.</li>`;
      $("checkBadge").textContent = "";
      $("checkBadge").className = "badge";
      $("form").innerHTML = `<p class="muted small">Reading documents… fields will appear here.</p>`;
      return;
    }
    renderChecks(emp);
    renderForm(emp);
  }

  function renderChecks(emp) {
    const ev = evaluate(emp.merged, emp.edits);
    const b = $("checkBadge");
    if (ev.errors) { b.textContent = `${ev.errors} to fix`; b.className = "badge error"; }
    else if (ev.warns) { b.textContent = `${ev.warns} to check`; b.className = "badge warn"; }
    else { b.textContent = "All clear"; b.className = "badge good"; }
    const list = ev.issues.length ? ev.issues
      : [{ level: "good", text: "Every required field was found, IDs match across documents and nothing is expired." }];
    $("issueList").innerHTML = list.map((i) =>
      `<li class="${i.level}" ${i.field ? `data-field="${i.field}"` : ""}>${i.level === "error" ? "●" : i.level === "warn" ? "▲" : "✓"} <span>${esc(i.text)}</span></li>`
    ).join("");
    return ev;
  }

  const GROUPS = [
    ["Identity", ["name_en", "name_ar", "emirates_id", "passport_no", "dob", "gender", "nationality", "blood_group", "marital_status"]],
    ["Job & residence", ["job_title", "residence_place", "religion", "sect", "education", "major"]],
    ["Certificates & validity", ["h2s_validity", "lifesaving_validity", "medical_validity", "induction_date", "certificate_no"]],
    ["Fieldglass", ["work_order", "fg_status", "loa_id", "location", "remarks"]],
    ["Contract & contacts", ["main_contract", "subcontractor", "email", "cell_phone", "emergency_phone"]],
  ];
  const FIELD_BY_KEY = Object.fromEntries(FIELDS.map((f) => [f.key, f]));

  function renderForm(emp) {
    const ev = evaluate(emp.merged, emp.edits);
    $("form").innerHTML = GROUPS.map(([title, keys]) => `
      <div class="form-group"><h4>${title}</h4><div class="fields">
        ${keys.map((k) => fieldHtml(FIELD_BY_KEY[k], ev.values[k], ev.meta[k] || {})).join("")}
      </div></div>`).join("");
  }
  function fieldHtml(f, v, m) {
    const id = `f_${f.key}`;
    let input;
    if (f.type === "select") {
      const opts = f.options.includes(v) || !v ? f.options : [v, ...f.options];
      input = `<select id="${id}" data-key="${f.key}">${opts.map((o) => `<option value="${esc(o)}" ${o === v ? "selected" : ""}>${esc(o || "—")}</option>`).join("")}${!v && !opts.includes("") ? `<option value="" selected>— select —</option>` : ""}</select>`;
    } else if (f.type === "date") {
      input = `<input id="${id}" data-key="${f.key}" type="date" value="${esc(v)}" />`;
    } else {
      input = `<input id="${id}" data-key="${f.key}" type="text" value="${esc(v)}" ${f.rtl ? 'dir="rtl" lang="ar"' : ""} ${f.type === "eid" ? 'inputmode="numeric"' : ""} />`;
    }
    const status = m.status || "ok";
    let src = "";
    if (m.edited) src = "Edited by you";
    else if (m.source === "Default") src = "Default — please confirm";
    else if (m.source) src = `From ${m.source}${m.confidence === "low" ? " · low confidence" : ""}`;
    else if (f.manual) src = "Not on the documents — enter manually";
    const conflict = m.conflict ? `<span class="conflict">Also seen: ${esc(m.conflict.join(" · "))}</span>` : "";
    return `<div class="field ${status}" id="wrap_${f.key}">
      <label for="${id}">${esc(f.label)} <span class="col">· ${f.col}</span></label>
      ${input}
      <span class="src" title="${esc(m.file || "")}">${esc(src)}</span>${conflict}
    </div>`;
  }

  // Field edits
  $("form").addEventListener("change", (e) => {
    const key = e.target.dataset.key;
    const emp = activeEmp();
    if (!key || !emp) return;
    emp.edits[key] = e.target.value.trim();
    if (key === "religion") {
      const cur = evaluate(emp.merged, emp.edits).values.sect;
      if (emp.edits.religion !== "Muslim") emp.edits.sect = "";
      else if (!cur) emp.edits.sect = "Sunni";
    }
    renderChecks(emp);
    renderForm(emp);
    renderSidebar();
    $("empTitle").textContent = empName(emp);
  });
  $("issueList").addEventListener("click", (e) => {
    const li = e.target.closest("li[data-field]");
    if (!li) return;
    const w = $(`wrap_${li.dataset.field}`);
    if (!w) return;
    w.scrollIntoView({ behavior: "smooth", block: "center" });
    w.classList.remove("flash"); void w.offsetWidth; w.classList.add("flash");
    const inp = w.querySelector("input,select");
    if (inp) setTimeout(() => inp.focus(), 300);
  });
  // Doc type change / retry
  $("docList").addEventListener("change", (e) => {
    const id = Number(e.target.dataset.doc);
    const emp = activeEmp();
    const d = emp && emp.files.find((x) => x.id === id);
    if (!d) return;
    d.typeOverride = e.target.value;
    remerge(emp);
    renderEmployee(emp);
    renderSidebar();
  });
  $("docList").addEventListener("click", (e) => {
    const id = Number(e.target.dataset.retry);
    if (!id) return;
    const emp = activeEmp();
    const d = emp && emp.files.find((x) => x.id === id);
    if (!d) return;
    d.status = "queued";
    d.error = "";
    renderEmployee(emp);
    pump();
  });
  $("addMoreBtn").addEventListener("click", () => $("moreInput").click());
  $("moreInput").addEventListener("change", async (e) => {
    const emp = activeEmp();
    const files = [...e.target.files].filter((f) => ACCEPT.test(f.name));
    e.target.value = "";
    if (!emp || !files.length) return;
    files.forEach((f) => emp.files.push(newDoc(f)));
    renderEmployee(emp);
    renderSidebar();
    pump();
  });

  /* ───────── write to Excel ───────── */
  function finalValues(emp) {
    return evaluate(emp.merged, emp.edits);
  }
  $("addRowBtn").addEventListener("click", () => {
    const emp = activeEmp();
    if (!emp) return;
    const ev = finalValues(emp);
    if (ev.errors && !window.confirm(`There ${ev.errors === 1 ? "is 1 problem" : `are ${ev.errors} problems`} still flagged. Add this employee anyway?`)) return;
    const res = FGExcel.writeEmployee(state.wb, ev.values, FIELDS);
    emp.row = res.row;
    emp.added = true;
    state.wbDirty = true;
    toast(res.existing ? `Updated existing row ${res.row} (same Emirates ID).` : `Added to row ${res.row}.`);
    render();
  });
  $("singleXlsxBtn").addEventListener("click", async () => {
    const emp = activeEmp();
    if (!emp) return;
    const ev = finalValues(emp);
    const r = await fetch("template.xlsx");
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(await r.arrayBuffer());
    FGExcel.writeEmployee(wb, ev.values, FIELDS);
    const safe = (ev.values.name_en || "Employee").replace(/[^\w\s-]/g, "").trim();
    await downloadWorkbook(wb, `Fieldglass - ${safe}.xlsx`);
  });

  /* ───────── boot ───────── */
  (async () => {
    state.token = store.get("fg_token");
    state.user = store.get("fg_user") || "";
    if (state.token) {
      try { await showApp(); return; } catch {}
    }
    showLogin();
  })();
})();
