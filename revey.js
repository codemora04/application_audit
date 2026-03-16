import { supabase } from "./supabase.js";
import { getCurrentAuditPeriod, getAuditPeriodStartDate, compressImage, showLoading, hideLoading, handleSupabaseError } from "./utils.js";

/* ===================== NAVIGATION PROTECTION ===================== */
window.addEventListener("beforeunload", (e) => {
  const zone = document.getElementById("zone")?.value;
  if (zone) {
    e.preventDefault();
    e.returnValue = "";
  }
});

const btnRetour = document.querySelector(".btn-retour-rouge");
if (btnRetour) {
  btnRetour.addEventListener("click", (e) => {
    const zone = document.getElementById("zone")?.value;
    if (zone && !confirm("Un audit est en cours. Voulez-vous vraiment quitter ? Votre progression locale sera conservée mais l'audit ne sera pas finalisé.")) {
      e.preventDefault();
    }
  });
}

/* ===================== SESSION CHECK ===================== */
const { data: sessionWrap } = await supabase.auth.getSession();
if (!sessionWrap?.session) window.location.href = "login.html";

/* ===================== LOAD DICT FROM SUPABASE ===================== */
async function loadDict(name) {
  const { data, error } = await supabase
    .from("app_data")
    .select("data")
    .eq("name", name)
    .single();

  if (error) {
    console.error("Cannot load", name, error.message);
    alert("Erreur: impossible de charger les données depuis Supabase.");
    return null;
  }
  return data?.data ?? null;
}

const DICT_REVEY = await loadDict("DICT_REVEY");
console.log("DICT_REVEY:", DICT_REVEY);
if (!DICT_REVEY) throw new Error("DICT_REVEY not loaded");

/* ===================== DOM ===================== */
const atelierSelect = document.getElementById("atelier");

const auditContainer = document.getElementById("audit-container");
const auditSelect = document.getElementById("audit");

const zoneContainer = document.getElementById("zone-container");
const zoneSelect = document.getElementById("zone");

const rubriqueContainer = document.getElementById("rubrique-container");
const rubriquesList = document.getElementById("rubriques-list");

const downloadBtn = document.getElementById("downloadPdf");
const downloadScoreBtn = document.getElementById("downloadScoreBtn");

const userId = (await supabase.auth.getUser()).data.user?.id;
const username = localStorage.getItem("username") || "";

/* ===================== DB STATE ===================== */
let currentSessionId = null;

/* ===================== PAGE STATE (persist across refresh) ===================== */
const STATE_KEY = "revey_state";

function saveState() {
  const state = {
    atelier: atelierSelect.value,
    audit: auditSelect.value,
    zone: zoneSelect.value,
  };
  localStorage.setItem(STATE_KEY, JSON.stringify(state));
}

function clearState() {
  localStorage.removeItem(STATE_KEY);
}

async function restoreState() {
  const raw = localStorage.getItem(STATE_KEY);
  if (!raw) return;

  let state;
  try { state = JSON.parse(raw); } catch { return; }

  const { atelier, audit, zone } = state;
  if (!atelier) return;

  // Restore atelier
  atelierSelect.value = atelier;
  if (!atelierSelect.value) return;

  // Rebuild audits
  const auditsObj = DICT_REVEY?.[atelier] || {};
  resetSelect(auditSelect);
  Object.keys(auditsObj).forEach((a) => {
    const opt = document.createElement("option");
    opt.value = a;
    opt.textContent = a;
    auditSelect.appendChild(opt);
  });
  auditContainer?.classList.remove("hidden");

  if (!audit) return;
  auditSelect.value = audit;
  if (!auditSelect.value) return;

  // Rebuild zones
  const zonesObj = DICT_REVEY?.[atelier]?.[audit] || {};
  resetSelect(zoneSelect);
  Object.keys(zonesObj).forEach((z) => {
    const opt = document.createElement("option");
    opt.value = z;
    opt.textContent = z;
    zoneSelect.appendChild(opt);
  });
  zoneContainer?.classList.remove("hidden");
  refreshSelectsProgress();

  if (!zone) return;
  zoneSelect.value = zone;
  if (!zoneSelect.value) return;

  const zoneData = DICT_REVEY?.[atelier]?.[audit]?.[zone];
  if (!zoneData) return;

  let rubriquesObj = zoneData;
  if (Array.isArray(zoneData)) {
    rubriquesObj = Object.fromEntries(zoneData.map((r) => [r, []]));
  }

  showRubriques(rubriquesObj);
  refreshSelectsProgress();
}

/* ===================== HELPERS ===================== */
function resetSelect(selectEl) {
  if (!selectEl) return;
  selectEl.innerHTML = `<option value="">--Choisir--</option>`;
}

function hideAllBelowAtelier() {
  auditContainer?.classList.add("hidden");
  zoneContainer?.classList.add("hidden");
  rubriqueContainer?.classList.add("hidden");
  downloadBtn?.classList.add("hidden");
  downloadScoreBtn?.classList.add("hidden");

  resetSelect(auditSelect);
  resetSelect(zoneSelect);
  if (rubriquesList) rubriquesList.innerHTML = "";
  document.getElementById("progress-root")?.classList.add("hidden");
  document.getElementById("summary-section")?.classList.add("hidden");
}

/* ===================== ROW COMPLETION + COLOR ===================== */
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
  if (!tr) return;
  if (isRowComplete(tr)) tr.classList.add("row-complete");
  else tr.classList.remove("row-complete");
}

/* ===================== PROGRESSION TRACKING ===================== */
function getProgressKey(atelier, audit, zone) {
  const period = getCurrentAuditPeriod(audit);
  return `prog_${username}_${period}_${atelier}_${audit}_${zone}`;
}

function updateProgressState(atelier, audit, zone, completedObj) {
  if (!atelier || !audit || !zone) return;
  const key = getProgressKey(atelier, audit, zone);
  localStorage.setItem(key, JSON.stringify(completedObj));
}

function getProgressState(atelier, audit, zone) {
  if (!atelier || !audit || !zone) return null;
  const key = getProgressKey(atelier, audit, zone);
  const data = localStorage.getItem(key);
  return data ? JSON.parse(data) : null;
}

function refreshSelectsProgress() {
  const atelier = atelierSelect.value;
  const audit = auditSelect.value;
  if (!atelier || !audit) return;

  const zoneOptions = zoneSelect.options;
  for (let i = 1; i < zoneOptions.length; i++) {
    const opt = zoneOptions[i];
    const zName = opt.value;

    const state = getProgressState(atelier, audit, zName);
    if (state && state.completed === state.total && state.total > 0) {
      if (!opt.textContent.startsWith("✅ ")) opt.textContent = "✅ " + zName;
    } else {
      opt.textContent = zName;
    }
  }
}

function computeViewProgress() {
  const atelier = atelierSelect.value;
  const audit = auditSelect.value;
  const zone = zoneSelect.value;
  if (!atelier || !audit || !zone) return;

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

  updateProgressState(atelier, audit, zone, { total: totalRows, completed: completedRows });

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

/* ===================== POPULATE ATELIERS ===================== */
resetSelect(atelierSelect);
Object.keys(DICT_REVEY || {}).forEach((atelier) => {
  const opt = document.createElement("option");
  opt.value = atelier;
  opt.textContent = atelier;
  atelierSelect.appendChild(opt);
});

atelierSelect.addEventListener("change", () => {
  hideAllBelowAtelier();
  clearState();

  const atelier = atelierSelect.value;
  if (!atelier) return;

  const auditsObj = DICT_REVEY?.[atelier] || {};
  const audits = Object.keys(auditsObj);

  resetSelect(auditSelect);
  audits.forEach((a) => {
    const opt = document.createElement("option");
    opt.value = a;
    opt.textContent = a;
    auditSelect.appendChild(opt);
  });

  auditContainer?.classList.remove("hidden");
  saveState();
});


auditSelect.addEventListener("change", async () => {
  zoneContainer?.classList.add("hidden");
  rubriqueContainer?.classList.add("hidden");
  downloadBtn?.classList.add("hidden");
  downloadScoreBtn?.classList.add("hidden");

  resetSelect(zoneSelect);
  if (rubriquesList) rubriquesList.innerHTML = "";

  const atelier = atelierSelect.value;
  const audit = auditSelect.value;
  if (!atelier || !audit) return;

  const zonesObj = DICT_REVEY?.[atelier]?.[audit] || {};
  const zones = Object.keys(zonesObj);

  zones.forEach((z) => {
    const opt = document.createElement("option");
    opt.value = z;
    opt.textContent = z;
    zoneSelect.appendChild(opt);
  });

  zoneContainer?.classList.remove("hidden");
  refreshSelectsProgress();
  saveState();

  // ✅ check for duplicate session
  showLoading("Vérification en cours...");
  try {
      const now = new Date();
      const periodStartDate = getAuditPeriodStartDate(audit);

      const fullAuditName = `${atelier} — ${audit}`;
      const { data: existing } = await supabase
          .from("audit_sessions")
          .select("id")
          .eq("user_id", userId)
          .eq("audit", fullAuditName)
          .gte("created_at", periodStartDate)
          .limit(1);

      if (existing && existing.length > 0) {
          if (!confirm("Un audit de ce type a déjà été effectué ou est en cours pour cette période. Voulez-vous en créer un nouveau ?")) {
              auditSelect.value = "";
              hideLoading();
              return;
          }
      }

      currentSessionId = await createAuditSession(atelier, audit);
  } catch (e) {
      handleSupabaseError(e, "Erreur initialisation session");
      return;
  } finally {
      hideLoading();
  }
});


zoneSelect.addEventListener("change", async () => {
  rubriqueContainer?.classList.add("hidden");
  downloadBtn?.classList.add("hidden");
  downloadScoreBtn?.classList.add("hidden");
  if (rubriquesList) rubriquesList.innerHTML = "";

  const atelier = atelierSelect.value;
  const audit = auditSelect.value;
  const zone = zoneSelect.value;
  if (!atelier || !audit || !zone) return;

  const zoneData = DICT_REVEY?.[atelier]?.[audit]?.[zone];
  if (!zoneData) return;

  // Si c'est un tableau ["R1","R2"], on le transforme en objet
  let rubriquesObj = zoneData;
  if (Array.isArray(zoneData)) {
    rubriquesObj = Object.fromEntries(zoneData.map((r) => [r, []]));
  }

  showRubriques(rubriquesObj);
  refreshSelectsProgress();
  saveState();

  if (currentSessionId) {
      await updateAuditSession({ zone });
  }
});

await restoreState();

/* ===================== DB HELPERS ===================== */
function safeUUID() {
  if (typeof crypto !== "undefined" && crypto.randomUUID) return crypto.randomUUID();
  return `${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

async function createAuditSession(atelier, audit) {
  const { data: session, error } = await supabase
    .from("audit_sessions")
    .insert({
      user_id: userId,
      audit: `${atelier} — ${audit}`,
      zone: null,
      souszone: null,
      created_at: new Date().toISOString(),
    })
    .select("id")
    .single();

  if (error) throw error;
  currentSessionId = session.id;
  return session.id;
}

async function updateAuditSession(patch) {
  if (!currentSessionId) return;
  await supabase.from("audit_sessions").update(patch).eq("id", currentSessionId);
}

async function uploadImageToStorage(file) {
  if (!file) return "";
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
function isSafetyAuditSelected() {
  // adapte le texte EXACT à ton option dans Supabase si besoin
  return auditSelect?.value === "Safety - Chasse aux anomalies";
}

function statusValueToLabel(value) {
  if (isSafetyAuditSelected()) {
    if (value === "oui") return "Oui";
    if (value === "non") return "Non";
    if (value === "na") return "Non applicable";
    return "";
  } else {
    if (value === "1") return "Good";
    if (value === "2") return "Acceptable";
    if (value === "3") return "Unsatisfactory";
    return "";
  }
}

/* ===================== SHOW RUBRIQUES ===================== */
function showRubriques(rubriquesObj) {
  if (!rubriquesList) return;
  rubriquesList.innerHTML = "";

  Object.entries(rubriquesObj || {}).forEach(([rubrique, questions], index) => {
    // Header (accordion)
    const header = document.createElement("div");
    header.className = "rubrique-header";
    header.innerHTML = `&#9654; ${rubrique}`;

    // Wrapper
    const tableWrapper = document.createElement("div");
    tableWrapper.className = "rubrique-table-wrapper hidden";

    // Table
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

      const safety = isSafetyAuditSelected();

      let statusOptions = `<option value="">--</option>`;
      if (safety) {
        statusOptions += `
          <option value="oui">Oui</option>
          <option value="non">Non</option>
          <option value="na">Non applicable</option>
        `;
      } else {
        statusOptions += `
          <option value="1">Good</option>
          <option value="2">Acceptable</option>
          <option value="3">Unsatisfactory</option>
        `;
      }

      tr.innerHTML = `
        <td>${q}</td>
        <td>
          <select name="status_${index}_${qIndex}">
            ${statusOptions}
          </select>
        </td>
        <td>
          <input type="text" name="comment_${index}_${qIndex}" placeholder="Commentaire..." />
        </td>
        <td>
          <input type="file" name="image_${index}_${qIndex}" accept="image/*" capture="environment" />
        </td>
      `;

      tbody.appendChild(tr);

      const statusEl = tr.querySelector("select");
      const commentEl = tr.querySelector('input[type="text"]');
      const fileEl = tr.querySelector('input[type="file"]');

      const rubriqueTitle = rubrique;
      const questionText = q;

      async function onRowChange() {
        const val = statusEl?.value || "";
        const statusLabel = statusValueToLabel(val);
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

      [statusEl, commentEl].forEach((el) => {
        if (!el) return;
        el.addEventListener("change", onRowChange);
        el.addEventListener("input", onRowChange);
      });
      fileEl.addEventListener("change", onRowChange);

      updateRowColor(tr);
    });

    tableWrapper.appendChild(table);

    // Accordion toggle
    header.addEventListener("click", () => {
      tableWrapper.classList.toggle("hidden");
      header.innerHTML =
        (tableWrapper.classList.contains("hidden") ? "&#9654;" : "&#9660;") + " " + rubrique;
    });

    rubriquesList.appendChild(header);
    rubriquesList.appendChild(tableWrapper);
  });

  rubriqueContainer?.classList.remove("hidden");
  downloadBtn?.classList.remove("hidden");
  downloadScoreBtn?.classList.remove("hidden");
}

/* ===================== IMAGE COMPRESS ===================== */
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

    const atelier = atelierSelect?.value || "";
    const audit = auditSelect?.value || "";
    const zone = zoneSelect?.value || "";
    const username = localStorage.getItem("username") || "Auditeur";
    const dateStr = new Date().toLocaleDateString("fr-FR");
    const timeStr = new Date().toLocaleTimeString("fr-FR");

    // Colors
    const colorSlate = [30, 41, 59]; // #1e293b
    const colorEmerald = [16, 185, 129]; // #e12020

    const addPageDesign = (pageNum) => {
      // --- En-tete ---
      doc.setFillColor(...colorSlate);
      doc.rect(0, 0, pageWidth, 40, "F");

      try {
        doc.addImage("logo2.png", "PNG", 14, 8, 35, 15);
      } catch (e) {
        doc.setTextColor(255, 255, 255);
        doc.setFontSize(16);
        doc.text("REVEY", 14, 18);
      }

      doc.setTextColor(255, 255, 255);
      doc.setFontSize(18);
      doc.setFont(undefined, "bold");
      doc.text("RAPPORT D'AUDIT", pageWidth - 14, 18, { align: "right" });

      doc.setFontSize(10);
      doc.setFont(undefined, "normal");
      doc.text(`${atelier} — ${audit}`, pageWidth - 14, 25, { align: "right" });
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
      doc.text(`Auditeur : ${username}`, 16, y);
      y += 10;

      // Score Summary Section
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
          const statusVal = tr.querySelector("select")?.value || "";
          const statusLabel = statusValueToLabel(statusVal);
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
            minCellHeight: 25,
          },
          alternateRowStyles: {
            fillColor: [248, 250, 252],
          },
          columnStyles: {
            0: { cellWidth: 70 },
            1: { cellWidth: 25, halign: "center" },
            2: { cellWidth: 55 },
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

      doc.save(`Rapport_Audit_Revey_${dateStr.replace(/\//g, "-")}.pdf`);
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
    const rows = wrapper?.querySelectorAll("tbody tr") || [];

    const totalQuestions = rows.length;
    let goodCount = 0;

    rows.forEach((tr) => {
      const status = tr.querySelector("select")?.value;

      // Good = "1" OU Oui = "oui"
      if (status === "1" || status === "oui") goodCount++;
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

/* ===================== PDF DOWNLOAD (SCORES TABLE) ===================== */
downloadScoreBtn?.addEventListener("click", () => {
  try {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF();

    const scores = computeScores();

    const audit = auditSelect?.value || "";
    const zone = zoneSelect?.value || "";

    const username = localStorage.getItem("username") || "";
    const date = new Date().toLocaleDateString("fr-FR");

    const rows = [
      ["Société", "REVEY"],
      ["Département / Audit", audit],
      ["Zone", zone],
      ["Date d'audit", date],
      ["Auditeur", username],
    ];

    let total = 0;
    const count = scores.length;

    scores.forEach((s) => {
      rows.push([s.rubrique, s.score]);
      total += parseFloat(s.score);
    });

    let totalFinal = 0;
    if (count > 0) totalFinal = total / count;

    totalFinal = Math.round(totalFinal * 10) / 10;
    rows.push(["Total", totalFinal + " %"]);

    doc.setFontSize(14);
    doc.text("Score Audit", 14, 15);

    doc.autoTable({
      startY: 25,
      body: rows,
      theme: "grid",

      styles: {
        fontSize: 10,
        cellPadding: 3,
        lineWidth: 0.2,
        lineColor: [180, 180, 180],
        valign: "middle",
        textColor: [25, 27, 35],
      },

      columnStyles: {
        0: { cellWidth: 90, fontStyle: "bold" },
        1: { cellWidth: 90, halign: "center" },
      },

      didParseCell: function (data) {
        // colonne gauche en vert clair
        if (data.section === "body" && data.column.index === 0) {
          data.cell.styles.fillColor = [198, 224, 180];
        }
        // Total en gris
        if (data.section === "body" && data.row.index === rows.length - 1) {
          data.cell.styles.fillColor = [220, 220, 220];
          data.cell.styles.fontStyle = "bold";
        }
      },
    });

    doc.save("score_audit_tableau.pdf");
  } catch (e) {
    console.error("Erreur Score PDF:", e);
    alert("Erreur Score PDF: " + e.message);
  }
});