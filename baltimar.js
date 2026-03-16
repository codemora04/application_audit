import { supabase } from "./supabase.js";
import { getCurrentAuditPeriod, getAuditPeriodStartDate, compressImage, showLoading, hideLoading, handleSupabaseError } from "./utils.js";

/* ===================== NAVIGATION PROTECTION ===================== */
window.addEventListener("beforeunload", (e) => {
  if (currentSessionId) {
    e.preventDefault();
    e.returnValue = "";
  }
});

const btnRetour = document.querySelector(".btn-retour");
if (btnRetour) {
  btnRetour.addEventListener("click", (e) => {
    if (currentSessionId && !confirm("Un audit est en cours. Voulez-vous vraiment quitter ? Votre progression locale sera conservée mais l'audit ne sera pas finalisé.")) {
      e.preventDefault();
    } else {
        currentSessionId = null; // Allow leaving
    }
  });
}

/* ===================== SESSION CHECK ===================== */
const { data } = await supabase.auth.getSession();
if (!data.session) window.location.href = "login.html";

const userId = data.session.user.id;

/* ===================== LOAD DICT FROM SUPABASE ===================== */
async function loadDict(name) {
  const { data, error } = await supabase
    .from("app_data")
    .select("data")
    .eq("name", name)
    .order("data", { ascending: true })
    .single();

  if (error) {
    console.error("Cannot load", name, error.message);
    alert("Erreur: impossible de charger les données depuis Supabase.");
    return null;
  }
  return data.data;
}

const DICT_BALTIMAR = await loadDict("DICT_BALTIMAR");
if (!DICT_BALTIMAR) throw new Error("DICT_BALTIMAR not loaded");

/* ===================== DOM ===================== */
const auditSelect = document.getElementById("audit");

const zoneContainer = document.getElementById("zone-container");
const zoneSelect = document.getElementById("zone");

const souszoneContainer = document.getElementById("souszone-container");
const souszoneSelect = document.getElementById("souszone");

const rubriqueContainer = document.getElementById("rubrique-container");
const rubriquesList = document.getElementById("rubriques-list");

const downloadBtn = document.getElementById("downloadPdf");
const downloadScoreBtn = document.getElementById("downloadScoreBtn");

const username = localStorage.getItem("username") || "";

/* ===================== DB STATE ===================== */
let currentSessionId = null;

/* ===================== PAGE STATE (persist across refresh) ===================== */
const STATE_KEY = "baltimar_state";

function saveState() {
  const state = {
    audit: auditSelect.value,
    zone: zoneSelect.value,
    souszone: souszoneSelect.value,
    sessionId: currentSessionId,
  };
  localStorage.setItem(STATE_KEY, JSON.stringify(state));
}

function clearState() {
  localStorage.removeItem(STATE_KEY);
}

/* Restore saved selections silently (no new session creation) */
async function restoreState() {
  const raw = localStorage.getItem(STATE_KEY);
  if (!raw) return;

  let state;
  try { state = JSON.parse(raw); } catch { return; }

  const { audit, zone, souszone, sessionId } = state;
  if (!audit) return;

  // Restore audit
  auditSelect.value = audit;
  if (!auditSelect.value) return; // value no longer in list

  currentSessionId = sessionId || null;

  // Rebuild zones
  const zonesObj = DICT_BALTIMAR[audit] || {};
  const zones = Object.keys(zonesObj);
  resetSelect(zoneSelect);
  zones.forEach((z) => {
    const opt = document.createElement("option");
    opt.value = z;
    opt.textContent = z;
    zoneSelect.appendChild(opt);
  });
  zoneContainer.classList.remove("hidden");
  refreshSelectsProgress();

  if (!zone) return;
  zoneSelect.value = zone;
  if (!zoneSelect.value) return;

  const zoneData = DICT_BALTIMAR[audit]?.[zone];
  if (!zoneData) return;

  const isGWP = audit === "Audit GWP-Agence" || audit === "Audit GWP-Usines";
  const values = typeof zoneData === "object" && zoneData !== null ? Object.values(zoneData) : [];
  const zoneIsDirectQuestions = Array.isArray(zoneData);
  const zoneIsDirectRubriques = values.length > 0 && values.every((v) => Array.isArray(v));

  if (isGWP || zoneIsDirectQuestions || zoneIsDirectRubriques) {
    showRubriques(zoneData);
    refreshSelectsProgress();
    return;
  }

  // Rebuild sous-zones
  const souszones = Object.keys(zoneData);
  resetSelect(souszoneSelect);
  souszones.forEach((sz) => {
    const opt = document.createElement("option");
    opt.value = sz;
    opt.textContent = sz;
    souszoneSelect.appendChild(opt);
  });
  souszoneContainer.classList.remove("hidden");
  refreshSelectsProgress();

  if (!souszone) return;
  souszoneSelect.value = souszone;
  if (!souszoneSelect.value) return;

  const rubriquesObj = DICT_BALTIMAR[audit]?.[zone]?.[souszone];
  if (!rubriquesObj) return;

  showRubriques(rubriquesObj);
}

/* ===================== HELPERS UI ===================== */
function resetSelect(selectEl) {
  selectEl.innerHTML = `<option value="">--Choisir--</option>`;
}

function hideAllBelowAudit() {
  zoneContainer.classList.add("hidden");
  souszoneContainer.classList.add("hidden");
  rubriqueContainer.classList.add("hidden");
  if (downloadBtn) downloadBtn.classList.add("hidden");
  if (downloadScoreBtn) downloadScoreBtn.classList.add("hidden");

  resetSelect(zoneSelect);
  resetSelect(souszoneSelect);
  rubriquesList.innerHTML = "";
  document.getElementById("progress-root")?.classList.add("hidden");
  document.getElementById("summary-section")?.classList.add("hidden");
}

/* ===================== ROW COMPLETION ===================== */
function isRowComplete(tr) {
  const statusEl = tr.querySelector("select");
  const commentEl = tr.querySelector('input[type="text"]');
  const fileEl = tr.querySelector('input[type="file"]');

  const status = statusEl?.value || "";
  const comment = commentEl?.value?.trim() || "";
  const hasFile = fileEl?.files && fileEl.files.length > 0;

  if (status === "") return false;
  if (comment === "") return false;

  // Image mandatory ONLY for: Acceptable(2), Unsatisfactory(3), Non(non)
  const imageMandatory = ["2", "3", "non"].includes(status);
  
  if (imageMandatory && !hasFile) return false;

  return true;
}

function updateRowColor(tr) {
  if (isRowComplete(tr)) tr.classList.add("row-complete");
  else tr.classList.remove("row-complete");
}

/* ===================== PROGRESSION TRACKING ===================== */
function getProgressKey(audit, zone, souszone) {
  const period = getCurrentAuditPeriod(audit);
  const id = `prog_${username}_${period}_${audit}_${zone}`;
  return souszone ? `${id}_${souszone}` : id;
}

function updateProgressState(audit, zone, souszone, completedObj) {
  if (!audit || !zone) return;
  const key = getProgressKey(audit, zone, souszone);
  localStorage.setItem(key, JSON.stringify(completedObj));
}

function getProgressState(audit, zone, souszone) {
  if (!audit || !zone) return null;
  const key = getProgressKey(audit, zone, souszone);
  const data = localStorage.getItem(key);
  return data ? JSON.parse(data) : null;
}

function checkSouszonesCompletion(audit, zone, souszonesList) {
  let allComplete = true;
  for (const sz of souszonesList) {
    const state = getProgressState(audit, zone, sz);
    if (!state || state.completed < state.total || state.total === 0) {
      allComplete = false;
      break;
    }
  }
  return allComplete;
}

function refreshSelectsProgress() {
  const audit = auditSelect.value;
  if (!audit) return;

  // Refresh Zones
  const zoneOptions = zoneSelect.options;
  const zonesObj = DICT_BALTIMAR[audit] || {};
  for (let i = 1; i < zoneOptions.length; i++) {
    const opt = zoneOptions[i];
    const zName = opt.value;

    // Check if it's direct or has sous-zones
    const zoneData = zonesObj[zName];
    const isDirect = Array.isArray(zoneData) || (typeof zoneData === "object" && zoneData !== null && Object.values(zoneData).every((v) => Array.isArray(v)));
    const isGWP = audit === "Audit GWP-Agence" || audit === "Audit GWP-Usines";

    let isComplete = false;

    if (isGWP || isDirect) {
      const state = getProgressState(audit, zName, null);
      if (state && state.completed === state.total && state.total > 0) isComplete = true;
    } else {
      const souszonesList = Object.keys(zoneData || {});
      isComplete = checkSouszonesCompletion(audit, zName, souszonesList);
    }

    if (isComplete) {
      if (!opt.textContent.startsWith("✅ ")) opt.textContent = "✅ " + zName;
    } else {
      opt.textContent = zName;
    }
  }

  // Refresh Sous-zones
  const zone = zoneSelect.value;
  if (!zone) return;
  const szOptions = souszoneSelect.options;
  for (let i = 1; i < szOptions.length; i++) {
    const opt = szOptions[i];
    const szName = opt.value;
    const state = getProgressState(audit, zone, szName);
    if (state && state.completed === state.total && state.total > 0) {
      if (!opt.textContent.startsWith("✅ ")) opt.textContent = "✅ " + szName;
    } else {
      opt.textContent = szName;
    }
  }
}

function computeViewProgress() {
  const audit = auditSelect.value;
  const zone = zoneSelect.value;
  const souszone = souszoneSelect.value;
  if (!audit || !zone) return;

  const tables = rubriquesList.querySelectorAll("table.questions-table");
  let totalRows = 0;
  let completedRows = 0;

  tables.forEach(table => {
    const rows = table.querySelectorAll("tbody tr");
    totalRows += rows.length;
    rows.forEach(tr => {
      if (isRowComplete(tr)) completedRows++;
    });
  });

  // Save progress for the current view
  updateProgressState(audit, zone, souszone || null, { total: totalRows, completed: completedRows });

  // Update Progress Bar
  const progressRoot = document.getElementById("progress-root");
  const progressFill = document.getElementById("progress-fill");
  const progressPercent = document.getElementById("progress-percent");

  if (progressRoot && totalRows > 0) {
    progressRoot.classList.remove("hidden");
    const percent = Math.round((completedRows / totalRows) * 100);
    progressFill.style.width = percent + "%";
    progressPercent.textContent = percent + "%";
  }

  // Update Summary if full
  if (totalRows > 0 && completedRows === totalRows) {
      showSummary();
  } else {
      document.getElementById("summary-section")?.classList.add("hidden");
  }

  // Update visuals in dropdowns
  refreshSelectsProgress();
}

function showSummary() {
    const summarySection = document.getElementById("summary-section");
    const summaryBody = document.getElementById("summary-body");
    if (!summarySection || !summaryBody) return;

    summaryBody.innerHTML = "";
    const rubriques = rubriquesList.querySelectorAll(".rubrique-header");
    let totalScore = 0;
    let count = 0;

    rubriques.forEach(header => {
        const title = header.textContent.replace(/[▶▼0-9]/g, "").trim();
        const wrapper = header.nextElementSibling;
        const rows = wrapper.querySelectorAll("tbody tr");
        let good = 0;
        rows.forEach(tr => {
            const val = tr.querySelector("select")?.value;
            if (val === "1" || val === "oui") good++;
        });
        const score = rows.length > 0 ? Math.round((good / rows.length) * 100) : 0;
        totalScore += score;
        count++;

        const tr = document.createElement("tr");
        tr.innerHTML = `<td>${title}</td><td style="text-align:right; font-weight:bold;">${score}%</td>`;
        summaryBody.appendChild(tr);
    });

    if (count > 0) {
        const avg = Math.round(totalScore / count);
        const trFinal = document.createElement("tr");
        trFinal.innerHTML = `<td>TOTAL GLOBAL</td><td style="text-align:right; font-weight:bold;">${avg}%</td>`;
        summaryBody.appendChild(trFinal);
    }

    summarySection.classList.remove("hidden");
}

/* ===================== STATUS OPTIONS ===================== */
function getStatusOptions(auditName) {
  if (auditName === "Audit Safety-Chasse au anomalies") {
    return `
      <option value="">--</option>
      <option value="oui">Oui</option>
      <option value="non">Non</option>
      <option value="na">Non applicable</option>
    `;
  }

  return `
    <option value="">--</option>
    <option value="1">Good</option>
    <option value="2">Acceptable</option>
    <option value="3">Unsatisfactory</option>
  `;
}

/* ===================== DB HELPERS ===================== */
function safeUUID() {
  if (typeof crypto !== "undefined" && crypto.randomUUID) return crypto.randomUUID();
  return `${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

async function createAuditSession(audit) {
  const { data: session, error } = await supabase
    .from("audit_sessions")
    .insert({
      user_id: userId,
      audit: audit,
      zone: null,
      souszone: null,
      created_at: new Date().toISOString(),
    })
    .select("id")
    .single();

  if (error) throw error;
  currentSessionId = session.id; // Correctly update currentSessionId
  return session.id;
}

async function updateAuditSession(patch) {
  if (!currentSessionId) return;
  await supabase.from("audit_sessions").update(patch).eq("id", currentSessionId);
}

async function uploadImageToStorage(file) {
  if (!file) return "";

  // bucket name must exist in Supabase Storage: audit-images
  const ext = (file.name.split(".").pop() || "jpg").toLowerCase();
  const filePath = `image_url/${safeUUID()}.${ext}`;

  const { error: upErr } = await supabase.storage
    .from("audit-images")
    .upload(filePath, file, { upsert: true });

  if (upErr) throw upErr;

  const { data } = supabase.storage.from("audit-images").getPublicUrl(filePath);
  return data.publicUrl || "";
}

async function saveAnswer({ rubriqueTitle, question, statusLabel, comment, file }) {
  if (!currentSessionId) return;

  let image_url = null;
  try {
    if (file) {
      const compressedFile = await compressImage(file);
      const url = await uploadImageToStorage(compressedFile);
      image_url = url || null;
    }
  } catch (e) {
    handleSupabaseError(e, "Erreur upload image");
    return;
  }

  const { error } = await supabase
    .from("audit_answers")
    .upsert(
      {
        session_id: currentSessionId,
        rubrique: rubriqueTitle,
        question: question,
        status: statusLabel,
        comment: comment,
        image_url: image_url,
        updated_at: new Date().toISOString(),
      },
      { onConflict: "session_id,rubrique,question" }
    );

  if (error) handleSupabaseError(error, "Erreur sauvegarde réponse");
}

/* ===================== INIT AUDITS SELECT ===================== */
Object.keys(DICT_BALTIMAR).forEach((audit) => {
  const opt = document.createElement("option");
  opt.value = audit;
  opt.textContent = audit;
  auditSelect.appendChild(opt);
});

/* ===================== EVENTS: AUDIT / ZONE / SOUSZONE ===================== */
auditSelect.addEventListener("change", async () => {
  hideAllBelowAudit();
  clearState();

  const audit = auditSelect.value;
  if (!audit) return;

  // ✅ check for duplicate session
  showLoading("Vérification en cours...");
  try {
      const now = new Date();
      const periodStartDate = getAuditPeriodStartDate(audit);

      const { data: existing } = await supabase
          .from("audit_sessions")
          .select("id")
          .eq("user_id", userId)
          .eq("audit", audit)
          .gte("created_at", periodStartDate)
          .limit(1);

      if (existing && existing.length > 0) {
          if (!confirm("Un audit de ce type a déjà été effectué ou est en cours pour cette période. Voulez-vous en créer un nouveau ?")) {
              auditSelect.value = "";
              hideLoading();
              return;
          }
      }

      currentSessionId = await createAuditSession(audit);
  } catch (e) {
      handleSupabaseError(e, "Erreur initialisation session");
      return;
  } finally {
      hideLoading();
  }

  const zonesObj = DICT_BALTIMAR[audit] || {};
  const zones = Object.keys(zonesObj);

  resetSelect(zoneSelect);
  zones.forEach((z) => {
    const opt = document.createElement("option");
    opt.value = z;
    opt.textContent = z;
    zoneSelect.appendChild(opt);
  });

  zoneContainer.classList.remove("hidden");
  refreshSelectsProgress();
  saveState();
});

zoneSelect.addEventListener("change", async () => {
  souszoneContainer.classList.add("hidden");
  rubriqueContainer.classList.add("hidden");
  if (downloadBtn) downloadBtn.classList.add("hidden");
  if (downloadScoreBtn) downloadScoreBtn.classList.add("hidden");

  resetSelect(souszoneSelect);
  rubriquesList.innerHTML = "";

  const audit = auditSelect.value;
  const zone = zoneSelect.value;
  if (!audit || !zone) return;

  // ✅ save zone in session
  await updateAuditSession({ zone, souszone: null });

  const zoneData = DICT_BALTIMAR[audit]?.[zone];
  if (!zoneData) return;

  // ✅ CAS GWP (déjà ton comportement)
  const isGWP = audit === "Audit GWP-Agence" || audit === "Audit GWP-Usines";
  if (isGWP) {
    showRubriques(zoneData);
    refreshSelectsProgress();
    saveState();
    return;
  }

  // ✅ NOUVEAU : CAS "pas de sous-zone"
  const values = typeof zoneData === "object" && zoneData !== null ? Object.values(zoneData) : [];

  const zoneIsDirectQuestions = Array.isArray(zoneData);
  const zoneIsDirectRubriques = values.length > 0 && values.every((v) => Array.isArray(v));

  if (zoneIsDirectQuestions || zoneIsDirectRubriques) {
    showRubriques(zoneData);
    refreshSelectsProgress();
    saveState();
    return;
  }

  // ✅ CAS NORMAL : il y a des sous-zones
  const souszones = Object.keys(zoneData);

  resetSelect(souszoneSelect);
  souszones.forEach((sz) => {
    const opt = document.createElement("option");
    opt.value = sz;
    opt.textContent = sz;
    souszoneSelect.appendChild(opt);
  });

  souszoneContainer.classList.remove("hidden");
  refreshSelectsProgress();
  saveState();
});

souszoneSelect.addEventListener("change", async () => {
  rubriqueContainer.classList.add("hidden");
  if (downloadBtn) downloadBtn.classList.add("hidden");
  if (downloadScoreBtn) downloadScoreBtn.classList.add("hidden");
  rubriquesList.innerHTML = "";

  const audit = auditSelect.value;
  const zone = zoneSelect.value;
  const souszone = souszoneSelect.value;
  if (!audit || !zone || !souszone) return;

  // ✅ save souszone in session
  await updateAuditSession({ souszone });

  const rubriquesObj = DICT_BALTIMAR[audit]?.[zone]?.[souszone];
  if (!rubriquesObj) return;

  showRubriques(rubriquesObj);
  saveState();
});

/* ---- Restore state on page load ---- */
await restoreState();

/* ===================== SHOW RUBRIQUES ===================== */
function showRubriques(rubriquesObj) {
  rubriquesList.innerHTML = "";

  // ✅ CAS 1 : directement un tableau de questions (pas de rubrique)
  if (Array.isArray(rubriquesObj)) {
    const header = document.createElement("div");
    header.className = "rubrique-header";
    header.innerHTML = "Questions";

    const tableWrapper = document.createElement("div");
    tableWrapper.className = "rubrique-table-wrapper";

    const table = document.createElement("table");
    table.className = "questions-table";

    table.innerHTML = `
      <thead>
        <tr>
          <th>Question</th>
          <th>Status</th>
          <th>Commentaire</th>
          <th>Image</th>
        </tr>
      </thead>
      <tbody></tbody>
    `;

    const tbody = table.querySelector("tbody");

    rubriquesObj.forEach((q, qIndex) => {
      const tr = document.createElement("tr");
      const auditName = auditSelect.value;

      tr.innerHTML = `
        <td>${q}</td>

        <td>
          <select name="status_0_${qIndex}">
            ${getStatusOptions(auditName)}
          </select>
        </td>

        <td>
          <input type="text" name="comment_0_${qIndex}" placeholder="Commentaire..." />
        </td>

        <td>
          <input type="file" name="image_0_${qIndex}" accept="image/*" />
        </td>
      `;

      tbody.appendChild(tr);

      const statusEl = tr.querySelector("select");
      const commentEl = tr.querySelector('input[type="text"]');
      const fileEl = tr.querySelector('input[type="file"]');

      const rubriqueTitle = "Questions";
      const questionText = q;

      async function onRowChange() {
        const statusLabel = statusEl?.selectedOptions?.[0]?.textContent?.trim() || "";
        const comment = commentEl?.value?.trim() || "";
        const file = fileEl?.files?.[0] || null;

        await saveAnswer({
          rubriqueTitle,
          question: questionText,
          statusLabel,
          comment,
          file,
        });

        updateRowColor(tr);
        computeViewProgress();
      }

      statusEl.addEventListener("change", onRowChange);
      commentEl.addEventListener("input", onRowChange);
      fileEl.addEventListener("change", onRowChange);

      updateRowColor(tr);
    });

    tableWrapper.appendChild(table);

    header.addEventListener("click", () => {
      tableWrapper.classList.toggle("hidden");
    });

    rubriquesList.appendChild(header);
    rubriquesList.appendChild(tableWrapper);
  }

  // ✅ CAS 2 : il y a des rubriques
  else {
    Object.entries(rubriquesObj).forEach(([rubrique, questions], index) => {
      const header = document.createElement("div");
      header.className = "rubrique-header";
      header.innerHTML = `&#9654; ${rubrique}`;

      const tableWrapper = document.createElement("div");
      tableWrapper.className = "rubrique-table-wrapper";

      const table = document.createElement("table");
      table.className = "questions-table";

      table.innerHTML = `
        <thead>
          <tr>
            <th>Question</th>
            <th>Status</th>
            <th>Commentaire</th>
            <th>Image</th>
          </tr>
        </thead>
        <tbody></tbody>
      `;

      const tbody = table.querySelector("tbody");

      (questions || []).forEach((q, qIndex) => {
        const tr = document.createElement("tr");
        const auditName = auditSelect.value;

        tr.innerHTML = `
          <td>${q}</td>

          <td>
            <select name="status_${index}_${qIndex}">
              ${getStatusOptions(auditName)}
            </select>
          </td>

          <td>
            <input type="text" name="comment_${index}_${qIndex}" placeholder="Commentaire..." />
          </td>

          <td>
            <input type="file" name="image_${index}_${qIndex}" accept="image/*" />
          </td>
        `;

        tbody.appendChild(tr);

        const statusEl = tr.querySelector("select");
        const commentEl = tr.querySelector('input[type="text"]');
        const fileEl = tr.querySelector('input[type="file"]');

        const rubriqueTitle = rubrique;
        const questionText = q;

        async function onRowChange() {
          const statusLabel = statusEl?.selectedOptions?.[0]?.textContent?.trim() || "";
          const comment = commentEl?.value?.trim() || "";
          const file = fileEl?.files?.[0] || null;

          await saveAnswer({
            rubriqueTitle,
            question: questionText,
            statusLabel,
            comment,
            file,
          });

          updateRowColor(tr);
          computeViewProgress();
        }

        statusEl.addEventListener("change", onRowChange);
        commentEl.addEventListener("input", onRowChange);
        fileEl.addEventListener("change", onRowChange);

        updateRowColor(tr);
      });

      tableWrapper.appendChild(table);

      header.addEventListener("click", () => {
        tableWrapper.classList.toggle("hidden");
        header.innerHTML =
          (tableWrapper.classList.contains("hidden") ? "&#9654;" : "&#9660;") +
          " " +
          rubrique;
      });

      tableWrapper.classList.add("hidden");

      rubriquesList.appendChild(header);
      rubriquesList.appendChild(tableWrapper);
    });
  }

  rubriqueContainer.classList.remove("hidden");
  if (downloadBtn) downloadBtn.classList.remove("hidden");
  if (downloadScoreBtn) downloadScoreBtn.classList.remove("hidden");
}

/* ===================== PDF: IMAGE COMPRESS ===================== */
async function readImageCompressed(file, maxW = 900, quality = 0.7) {
  if (!file || !file.type.startsWith("image/")) return "";

  const dataUrl = await new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => resolve(r.result);
    r.onerror = reject;
    r.readAsDataURL(file);
  });

  const img = await new Promise((resolve, reject) => {
    const im = new Image();
    im.onload = () => resolve(im);
    im.onerror = () => reject(new Error("Image illisible"));
    im.src = dataUrl;
  });

  const ratio = Math.min(1, maxW / img.width);
  const w = Math.round(img.width * ratio);
  const h = Math.round(img.height * ratio);

  const canvas = document.createElement("canvas");
  canvas.width = w;
  canvas.height = h;

  const ctx = canvas.getContext("2d");
  ctx.drawImage(img, 0, 0, w, h);

  return canvas.toDataURL("image/jpeg", quality);
}

/* ===================== PDF DOWNLOAD (FULL REPORT) ===================== */
downloadBtn?.addEventListener("click", async () => {
  try {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF();
    const pageWidth = doc.internal.pageSize.getWidth();
    const pageHeight = doc.internal.pageSize.getHeight();

    const audit = auditSelect?.value || "";
    const zone = zoneSelect?.value || "";
    const souszone = souszoneSelect?.value || "";
    const username = localStorage.getItem("username") || "Auditeur";
    const dateStr = new Date().toLocaleDateString("fr-FR");
    const timeStr = new Date().toLocaleTimeString("fr-FR");

    // Couleurs
    const colorSlate = [30, 41, 59]; // #1e293b
    const colorEmerald = [16, 185, 129]; // #10b981

    const addPageDesign = (pageNum) => {
      // --- En-tête ---
      doc.setFillColor(...colorSlate);
      doc.rect(0, 0, pageWidth, 40, "F");

      try {
        doc.addImage("logo1.png", "PNG", 14, 8, 35, 15);
      } catch (e) {
        doc.setTextColor(255, 255, 255);
        doc.setFontSize(16);
        doc.text("BALTIMAR", 14, 18);
      }

      doc.setTextColor(255, 255, 255);
      doc.setFontSize(18);
      doc.setFont(undefined, "bold");
      doc.text("RAPPORT D'AUDIT", pageWidth - 14, 18, { align: "right" });

      doc.setFontSize(10);
      doc.setFont(undefined, "normal");
      doc.text(`${audit}`, pageWidth - 14, 25, { align: "right" });
      doc.text(`Page ${pageNum}`, pageWidth - 14, 32, { align: "right" });

      // --- Pied de page ---
      doc.setDrawColor(226, 232, 240);
      doc.setLineWidth(0.1);
      doc.line(14, pageHeight - 15, pageWidth - 14, pageHeight - 15);

      doc.setFontSize(9);
      doc.setTextColor(51, 65, 85);
      doc.setFont(undefined, "normal");
      doc.text(`Généré le ${dateStr} à ${timeStr} par ${username}`, 14, pageHeight - 10);
      doc.text(`Page ${pageNum}`, pageWidth - 14, pageHeight - 10, { align: "right" });
    };

    const generatePDF = async () => {
      let currentPage = 1;
      addPageDesign(currentPage);

      // Info Section
      let y = 50;
      doc.setTextColor(...colorSlate);
      doc.setFontSize(12);
      doc.setFont(undefined, "bold");
      doc.text("Informations générales", 14, y);
      y += 8;

      doc.setFontSize(10);
      doc.setFont(undefined, "normal");
      doc.text(`Zone : ${zone}`, 16, y);
      y += 6;
      if (souszone) {
        doc.text(`Sous-zone : ${souszone}`, 16, y);
        y += 6;
      }
      doc.text(`Auditeur : ${username}`, 16, y);
      y += 10;

      // section résumé des scores
      const scores = computeScores();
      let totalGood = 0;
      let totalQ = 0;
      scores.forEach(s => {
        const val = parseFloat(s.score);
        totalGood += (val / 100);
        totalQ++;
      });
      const globalScore = totalQ > 0 ? Math.round((totalGood / totalQ) * 100) : 0;

      // Score Box
      doc.setDrawColor(...colorEmerald);
      doc.setLineWidth(0.5);
      doc.setFillColor(240, 253, 244);
      doc.roundedRect(14, y, pageWidth - 28, 20, 3, 3, "FD");

      doc.setTextColor(...colorEmerald);
      doc.setFontSize(11);
      doc.setFont(undefined, "bold");
      doc.text("SCORE GLOBAL", 25, y + 12);

      doc.setFontSize(16);
      doc.text(`${globalScore}%`, pageWidth - 40, y + 13);
      y += 30;

      // Rubriques & Tables
      const children = Array.from(rubriquesList.children);

      for (let i = 0; i < children.length; i++) {
        const el = children[i];
        if (!el.classList.contains("rubrique-header")) continue;

        const rubriqueTitle = el.textContent.replace("▼", "").replace("▶", "").trim();
        const wrapper = children[i + 1];
        const table = wrapper?.querySelector("table.questions-table");
        if (!table) continue;

        const rows = [];
        const imagesMap = new Map();
        const trs = Array.from(table.querySelectorAll("tbody tr"));

        for (let r = 0; r < trs.length; r++) {
          const tr = trs[r];
          const question = tr.children[0]?.textContent?.trim() || "";
          const statusSelect = tr.querySelector("select");
          const statusLabel = statusSelect?.selectedOptions?.[0]?.textContent?.trim() || "";
          const comment = tr.querySelector('input[type="text"]')?.value?.trim() || "";
          const fileInput = tr.querySelector('input[type="file"]');
          const file = fileInput?.files?.[0];

          if (file) {
            const imgData = await readImageCompressed(file, 900, 0.7);
            if (imgData) imagesMap.set(r, imgData);
          }
          rows.push([question, statusLabel, comment, ""]);
        }

        if (y > pageHeight - 60) {
          doc.addPage();
          currentPage++;
          addPageDesign(currentPage);
          y = 50;
        }

        doc.setTextColor(...colorSlate);
        doc.setFontSize(12);
        doc.setFont(undefined, "bold");
        doc.text(rubriqueTitle, 14, y);
        y += 6;

        doc.autoTable({
          startY: y,
          head: [["Question", "Status", "Commentaire", "Image"]],
          body: rows,
          margin: { top: 45, bottom: 20 },
          didDrawPage: (data) => {
            // Si autoTable crée une page, on dessine le design
            if (doc.internal.getNumberOfPages() > currentPage) {
              currentPage++;
              addPageDesign(currentPage);
            }
          },
          headStyles: {
            fillColor: colorEmerald,
            textColor: 255,
            fontStyle: "bold",
            halign: "center",
          },
          styles: {
            fontSize: 9,
            cellPadding: 3,
            valign: "middle",
            lineWidth: 0.1,
            lineColor: [200, 200, 200],
          },
          bodyStyles: {
            minCellHeight: 25, // Donner de l'espace pour l'image
          },
          alternateRowStyles: {
            fillColor: [248, 250, 252],
          },
          columnStyles: {
            0: { cellWidth: 70 },
            1: { cellWidth: 25, halign: "center" },
            2: { cellWidth: 60 },
            3: { cellWidth: 30 },
          },
          didParseCell: (data) => {
            if (data.section === "body" && data.column.index === 3) data.cell.text = [""];
            if (data.section === "body" && data.column.index === 1) {
              const val = data.cell.raw;
              if (val === "Good" || val === "Oui") data.cell.styles.textColor = [16, 185, 129];
              if (val === "Unsatisfactory" || val === "Non") data.cell.styles.textColor = [239, 68, 68];
            }
          },
          didDrawCell: (data) => {
            if (data.section === "body" && data.column.index === 3) {
              const imgData = imagesMap.get(data.row.index);
              if (imgData) {
                const x = data.cell.x + 2;
                const yImg = data.cell.y + 2;
                const w = data.cell.width - 4;
                const h = data.cell.height - 4;
                doc.addImage(imgData, "JPEG", x, yImg, w, h);
              }
            }
          },
        });

        y = doc.lastAutoTable.finalY + 15;
      }

      doc.save(`Rapport_Audit_Baltimar_${dateStr.replace(/\//g, "-")}.pdf`);
    };

    await generatePDF();
  } catch (e) {
    console.error("Erreur PDF:", e);
    alert("Erreur PDF: " + e.message);
  }
});

/* ===================== SCORE COMPUTE ===================== */
function computeScores() {
  const results = [];
  const headers = rubriquesList.querySelectorAll(".rubrique-header");

  headers.forEach((header, index) => {
    const rubriqueName = header.textContent.replace("▼", "").replace("▶", "").trim();

    const wrapper = header.nextElementSibling;
    const rows = wrapper.querySelectorAll("tbody tr");

    const totalQuestions = rows.length;
    let goodCount = 0;

    rows.forEach((tr) => {
      const status = tr.querySelector("select")?.value;

      // "Good" (value="1") OU "Oui" (value="oui")
      if (status === "1" || status === "oui") {
        goodCount++;
      }
    });

    let score = 0;
    if (totalQuestions > 0) score = (goodCount / totalQuestions) * 100;

    score = Math.round(score * 10) / 10;

    results.push({
      rubrique: rubriqueName || `S${index + 1}`,
      score: score + " %",
    });
  });

  return results;
}

/* ===================== EXCEL DOWNLOAD - SCORES COMPLETS (toutes zones/sous-zones) ===================== */
downloadScoreBtn?.addEventListener("click", async () => {
  const audit = auditSelect.value || "";
  const usernameVal = localStorage.getItem("username") || "";
  const date = new Date().toLocaleDateString("fr-FR");

  if (!audit) {
    alert("Veuillez d'abord sélectionner un type d'audit.");
    return;
  }

  // 1. Récupérer toutes les sessions de cet utilisateur pour cet audit
  const { data: sessions, error: sessErr } = await supabase
    .from("audit_sessions")
    .select("id, zone, souszone")
    .eq("audit", audit)
    .eq("user_id", userId);

  if (sessErr || !sessions?.length) {
    alert("Aucune donnée d'audit trouvée pour cet audit.");
    return;
  }

  const sessionIds = sessions.map((s) => s.id);

  // 2. Récupérer toutes les réponses pour ces sessions
  const { data: answers, error: ansErr } = await supabase
    .from("audit_answers")
    .select("session_id, rubrique, question, status")
    .in("session_id", sessionIds);

  if (ansErr) {
    alert("Erreur lors de la récupération des réponses.");
    return;
  }

  // 3. Construire un index session_id -> {zone, souszone}
  const sessionMap = {};
  sessions.forEach((s) => { sessionMap[s.id] = { zone: s.zone, souszone: s.souszone }; });

  // 4. Agréger les scores par zone et sous-zone
  // Structure: { zone: { souszone|"_direct": { rubrique: { good, total } } } }
  const tree = {};

  (answers || []).forEach((a) => {
    const sess = sessionMap[a.session_id];
    if (!sess?.zone) return;

    const zone = sess.zone;
    const souszone = sess.souszone || "_direct";
    const rubrique = a.rubrique || "Questions";
    const isGood = a.status === "Good" || a.status === "Oui";

    if (!tree[zone]) tree[zone] = {};
    if (!tree[zone][souszone]) tree[zone][souszone] = {};
    if (!tree[zone][souszone][rubrique]) tree[zone][souszone][rubrique] = { good: 0, total: 0 };

    tree[zone][souszone][rubrique].total++;
    if (isGood) tree[zone][souszone][rubrique].good++;
  });

  // 5. Construire le fichier Excel
  const wb = XLSX.utils.book_new();

  // --- Feuille RÉSUMÉ ---
  const summaryData = [
    ["Société", "BALTIMAR"],
    ["Département / Audit", audit],
    ["Date d'audit", date],
    ["Auditeur", usernameVal],
    [],
    ["Zone", "Sous-zone", "Score Zone (%)", "Score Sous-zone (%)"],
  ];

  const zoneScores = {}; // zone -> {good, total}

  Object.entries(tree).forEach(([zone, souszones]) => {
    let zoneGood = 0, zoneTotal = 0;

    Object.entries(souszones).forEach(([souszone, rubriques]) => {
      let szGood = 0, szTotal = 0;

      Object.values(rubriques).forEach(({ good, total }) => {
        szGood += good;
        szTotal += total;
        zoneGood += good;
        zoneTotal += total;
      });

      const szScore = szTotal > 0 ? Math.round((szGood / szTotal) * 1000) / 10 : 0;
      const szLabel = souszone === "_direct" ? "—" : souszone;
      summaryData.push([zone, szLabel, "", szScore]);
    });

    const zScore = zoneTotal > 0 ? Math.round((zoneGood / zoneTotal) * 1000) / 10 : 0;
    zoneScores[zone] = zScore;

    // Mettre le score de zone sur la première ligne de cette zone
    const firstIdx = summaryData.findLastIndex((r) => r[0] === zone);
    if (firstIdx >= 0) summaryData[firstIdx][2] = zScore;
  });

  const allGood = Object.values(tree).reduce((acc, szs) =>
    acc + Object.values(szs).reduce((a, rs) =>
      a + Object.values(rs).reduce((b, r) => b + r.good, 0), 0), 0);
  const allTotal = Object.values(tree).reduce((acc, szs) =>
    acc + Object.values(szs).reduce((a, rs) =>
      a + Object.values(rs).reduce((b, r) => b + r.total, 0), 0), 0);
  const globalScore = allTotal > 0 ? Math.round((allGood / allTotal) * 1000) / 10 : 0;

  summaryData.push([]);
  summaryData.push(["SCORE GLOBAL", "", globalScore, ""]);

  const wsSummary = XLSX.utils.aoa_to_sheet(summaryData);
  wsSummary["!cols"] = [{ wch: 30 }, { wch: 25 }, { wch: 18 }, { wch: 20 }];
  XLSX.utils.book_append_sheet(wb, wsSummary, "Résumé");

  // --- Une feuille par Zone ---
  Object.entries(tree).forEach(([zone, souszones]) => {
    const wsData = [
      ["Zone", zone],
      ["Score Zone (%)", zoneScores[zone]],
      [],
    ];

    Object.entries(souszones).forEach(([souszone, rubriques]) => {
      let szGood = 0, szTotal = 0;

      if (souszone !== "_direct") {
        wsData.push([`Sous-zone : ${souszone}`, ""]);
      }

      wsData.push(["Rubrique", "Score (%)", "Nb Good", "Nb Total"]);

      Object.entries(rubriques).forEach(([rubrique, { good, total }]) => {
        const score = total > 0 ? Math.round((good / total) * 1000) / 10 : 0;
        wsData.push([rubrique, score, good, total]);
        szGood += good;
        szTotal += total;
      });

      if (souszone !== "_direct") {
        const szScore = szTotal > 0 ? Math.round((szGood / szTotal) * 1000) / 10 : 0;
        wsData.push(["Total Sous-zone", szScore, szGood, szTotal]);
      }

      wsData.push([]);
    });

    const ws = XLSX.utils.aoa_to_sheet(wsData);
    ws["!cols"] = [{ wch: 35 }, { wch: 15 }, { wch: 12 }, { wch: 12 }];

    // Sanitize sheet name (max 31 chars, no special chars)
    const sheetName = zone.replace(/[:\\\/\?\*\[\]]/g, "").substring(0, 31);
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
  });

  const filename = `score_complet_baltimar_${date.replace(/\//g, "-")}.xlsx`;
  XLSX.writeFile(wb, filename);
});