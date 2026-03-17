import { supabase } from "./supabase.js";
import { showLoading, hideLoading, handleSupabaseError, clearAuditProgress } from "./utils.js";

/* ===================== SESSION CHECK & SECURITY ===================== */
const { data: sessionData } = await supabase.auth.getSession();
if (!sessionData.session) window.location.href = "login.html";

const savedRole = sessionStorage.getItem("audit_user_role");
if (savedRole !== "responsable") {
    window.location.href = "accueil.html";
}

const operatorNameStr = sessionStorage.getItem("audit_username") || "Responsable";
document.getElementById("operatorName").textContent = "Bienvenue, " + operatorNameStr;

/* ===================== DOM ===================== */
const companySelect = document.getElementById("companySelect");
const atelierContainer = document.getElementById("atelierContainer");
const atelierSelect = document.getElementById("atelierSelect");
const auditSelect = document.getElementById("auditSelect");
const chartContainer = document.getElementById("chartContainer");
const noDataMessage = document.getElementById("noDataMessage");
const logoutBtn = document.getElementById("logoutBtn");
const downloadBtn = document.getElementById("downloadBtn");
const canvas = document.getElementById("scoreChart");
const ctx = canvas.getContext("2d");
const zoneSpeechBubble = document.getElementById("zoneSpeechBubble");

let scoreChart = null;
let DICT_BALTIMAR = null;
let DICT_REVEY = null;



/* ===================== SAFETY AUDIT STRUCTURE ===================== */
const SAFETY_STRUCTURE = {
    "Accès Piétons": ["Voie d'accès", "Sol", "Signalétique", "Éclairage"],
    "Accès Engins": ["Voie d'accès", "Sol", "Signalétique", "Éclairage"],
    "Stockage Extérieur": ["Conformité", "Propreté", "Organisation"],
    "Parking": ["Marquage", "Signalétique", "État général"],
    "Zone de Déchargement": ["Sécurité", "Organisation", "Propreté"],
    "Vestiaires": ["Propreté", "Équipements", "Ventilation"],
    "Sanitaires": ["Propreté", "État", "Accessibilité"],
    "Réfectoire": ["Hygiène", "Équipements", "Propreté"],
    "Bureaux": ["Ergonomie", "Ordre", "Sécurité"],
    "Ateliers": ["Machines", "Propreté", "Rangement", "EPI", "Ventilation"],
    "Stockage Intérieur": ["Organisation", "Rayonnages", "Signalétique"],
    "Circulation Intérieure": ["Marquage", "Dégagements", "Signalétique"],
    "Issues de Secours": ["Accessibilité", "Signalétique", "Fonctionnement"],
    "Extincteurs": ["Présence", "Accessibilité", "Contrôle"],
    "Électricité": ["Armoires", "Câblage", "Conformité"]
};

/* ===================== LOAD DICT FROM SUPABASE ===================== */
async function loadDict(name) {
    const { data, error } = await supabase
        .from("app_data")
        .select("data")
        .eq("name", name)
        .single();

    if (error) {
        console.error("Cannot load dict", name, error);
        return null;
    }
    return data?.data ?? null;
}

/* ===================== RESET HELPERS ===================== */
function resetAuditSelect(placeholder = "-- Choisir un audit --") {
    auditSelect.innerHTML = `<option value="">${placeholder}</option>`;
    auditSelect.disabled = true;
}

function hideChart() {
    chartContainer.classList.add("hidden");
    noDataMessage.classList.add("hidden");
    const tableContainer = document.getElementById("summaryTableContainer");
    if (tableContainer) tableContainer.classList.add("hidden");
    if (zoneSpeechBubble) zoneSpeechBubble.classList.add("hidden");
}


/* ===================== POPULATE AUDITS ===================== */
// dict = the object whose keys are audit names
function populateAudits(dict) {
    if (!dict) {
        resetAuditSelect("Erreur de chargement");
        return;
    }
    auditSelect.innerHTML = `<option value="">-- Choisir un audit --</option>`;
    Object.keys(dict).forEach(audit => {
        const opt = document.createElement("option");
        opt.value = audit;
        opt.textContent = audit;
        auditSelect.appendChild(opt);
    });
    auditSelect.disabled = false;
}

/* ===================== COMPANY FILTER ===================== */
companySelect.addEventListener("change", async () => {
    const company = companySelect.value;

    // Reset everything below
    atelierContainer.style.display = "none";
    atelierSelect.innerHTML = `<option value="">-- Choisir un atelier --</option>`;
    resetAuditSelect("-- Sélectionner une entreprise d'abord --");
    hideChart();

    if (!company) return;

    if (company === "BALTIMAR") {
        // Load Baltimar dict once
        if (!DICT_BALTIMAR) DICT_BALTIMAR = await loadDict("DICT_BALTIMAR");
        populateAudits(DICT_BALTIMAR);

    } else if (company === "REVEY") {
        // Load Revey dict once
        if (!DICT_REVEY) DICT_REVEY = await loadDict("DICT_REVEY");

        if (DICT_REVEY) {
            // Populate ateliers from top-level keys of DICT_REVEY
            Object.keys(DICT_REVEY).forEach(atelier => {
                const opt = document.createElement("option");
                opt.value = atelier;
                opt.textContent = atelier;
                atelierSelect.appendChild(opt);
            });
            atelierContainer.style.display = "block";
            resetAuditSelect("-- Choisir un atelier d'abord --");
        } else {
            resetAuditSelect("Erreur de chargement de Revey");
        }
    }
});

/* ===================== ATELIER FILTER (Revey only) ===================== */
atelierSelect.addEventListener("change", () => {
    const atelier = atelierSelect.value;
    resetAuditSelect("-- Choisir un audit --");
    hideChart();

    if (!atelier || !DICT_REVEY) return;

    // DICT_REVEY[atelier] = { auditName: { zone: [...] } }
    populateAudits(DICT_REVEY[atelier]);
});


/* ===================== AUDIT FILTER & CHART ===================== */
auditSelect.addEventListener("change", async () => {
    const audit = auditSelect.value;
    hideChart();
    if (!audit) return;

    // 1. Fetch all sessions for this audit
    const { data: sessions, error: sessErr } = await supabase
        .from("audit_sessions")
        .select("id, zone, created_at")
        .eq("audit", audit)
        .order("created_at", { ascending: true });

    if (sessErr || !sessions || sessions.length === 0) {
        noDataMessage.classList.remove("hidden");
        return;
    }

    // If Revey is selected, filter sessions by atelier (zone prefix or via separate field)
    // The atelier filter is handled at the dict level (only audits from that atelier are shown)
    // so no extra filtering needed here.

    const sessionIds = sessions.map(s => s.id);

    // 2. Fetch answers
    const { data: answers, error: ansErr } = await supabase
        .from("audit_answers")
        .select("session_id, status")
        .in("session_id", sessionIds);

    if (ansErr || !answers || answers.length === 0) {
        noDataMessage.classList.remove("hidden");
        return;
    }

    // 3. Map sessions → zone + formatted date
    const sessionMap = {};
    const orderedLabels = [];

    sessions.forEach(s => {
        if (s.zone && s.created_at) {
            const dateStr = new Date(s.created_at).toLocaleDateString("fr-FR", {
                month: "short",
                year: "numeric"
            });
            sessionMap[s.id] = { zone: s.zone, date: dateStr };
            if (!orderedLabels.includes(dateStr)) orderedLabels.push(dateStr);
        }
    });

    // 4. Aggregate by date & zone
    const timelineStats = {};
    const uniqueZones = new Set();

    answers.forEach(ans => {
        const info = sessionMap[ans.session_id];
        if (!info) return;
        const { zone, date } = info;
        uniqueZones.add(zone);
        if (!timelineStats[date]) timelineStats[date] = {};
        if (!timelineStats[date][zone]) timelineStats[date][zone] = { total: 0, good: 0 };

        const isGood = ans.status === "Good" || ans.status === "Oui" || ans.status === "1";
        timelineStats[date][zone].total++;
        if (isGood) timelineStats[date][zone].good++;
    });

    const labels = orderedLabels;
    const zonesToChart = Array.from(uniqueZones);

    if (labels.length === 0 || zonesToChart.length === 0) {
        noDataMessage.classList.remove("hidden");
        return;
    }

    // 5. Build Chart.js datasets (one per zone)
    const datasets = zonesToChart.map((zone, index) => {
        const hue = (index * 137.5) % 360;
        const color = `hsl(${hue}, 70%, 50%)`;

        const dataPoints = labels.map(date => {
            if (timelineStats[date]?.[zone]) {
                const { good, total } = timelineStats[date][zone];
                return Math.round((good / total) * 100);
            }
            return null;
        });

        return {
            label: zone,
            data: dataPoints,
            borderColor: color,
            backgroundColor: color,
            borderWidth: 2,
            tension: 0.3,
            spanGaps: true
        };
    });

    renderChart(labels, datasets);
    renderSummaryTable(labels, timelineStats, zonesToChart);
    chartContainer.classList.remove("hidden");
    const tableContainer = document.getElementById("summaryTableContainer");
    if (tableContainer) tableContainer.classList.remove("hidden");
});



function renderSummaryTable(labels, timelineStats, zonesToChart) {
    const tableBody = document.getElementById("summaryTableBody");
    if (!tableBody) return;
    tableBody.innerHTML = "";

    // Show the most recent period first
    const lastDate = labels[labels.length - 1];

    zonesToChart.forEach(zone => {
        const stats = timelineStats[lastDate]?.[zone];
        const score = stats ? Math.round((stats.good / stats.total) * 100) : 0;

        const tr = document.createElement("tr");
        tr.innerHTML = `
            <td style="text-align:left; font-weight:500;">${zone}</td>
            <td style="text-align:center; font-weight:bold; color: var(--primary-color);">${score}%</td>
        `;
        tableBody.appendChild(tr);
    });
}

/* ===================== HIDE BUBBLE ON OUTSIDE CLICK ===================== */
document.addEventListener("click", (e) => {
    if (e.target.id !== "scoreChart" && zoneSpeechBubble) {
        zoneSpeechBubble.classList.add("hidden");
    }
});

/* ===================== RENDER CHART ===================== */
function renderChart(labels, datasets) {
    if (scoreChart) scoreChart.destroy();

    Chart.defaults.color = "#94a3b8";

    scoreChart = new Chart(ctx, {
        type: "line",
        data: { labels, datasets },
        options: {
            responsive: true,
            scales: {
                y: {
                    beginAtZero: true,
                    max: 100,
                    grid: { color: "rgba(255,255,255,0.1)" },
                    ticks: { callback: value => value + "%" }
                },
                x: {
                    grid: { color: "rgba(255,255,255,0.05)" }
                }
            },
            plugins: {
                legend: {
                    display: true,
                    position: "bottom",
                    labels: { color: "var(--text-main)" }
                },
                tooltip: {
                    callbacks: {
                        label: c => c.dataset.label + ": " + c.parsed.y + "%"
                    }
                }
            },
            onClick: (e, elements, chart) => {
                if (!zoneSpeechBubble) return;

                if (!elements || elements.length === 0) {
                    zoneSpeechBubble.classList.add("hidden");
                    return;
                }

                const datasetIndex = elements[0].datasetIndex;
                const zone = chart.data.datasets[datasetIndex].label;

                const auditName = auditSelect.value || "";
                if (!auditName.toLowerCase().includes("safety")) {
                    zoneSpeechBubble.classList.add("hidden");
                    return;
                }

                const subzones = SAFETY_STRUCTURE[zone];
                if (!subzones || subzones.length === 0) {
                    zoneSpeechBubble.classList.add("hidden");
                    return;
                }

                let html = `<div class="bubble-title">${zone}</div><ul class="bubble-list">`;
                subzones.forEach(sz => {
                    html += `<li>${sz}</li>`;
                });
                html += `</ul>`;
                zoneSpeechBubble.innerHTML = html;

                zoneSpeechBubble.classList.remove("hidden");
                zoneSpeechBubble.style.left = "-9999px";

                setTimeout(() => {
                    const x = e.native.pageX;
                    const y = e.native.pageY;
                    zoneSpeechBubble.style.left = (x - zoneSpeechBubble.offsetWidth / 2) + "px";
                    zoneSpeechBubble.style.top = (y - zoneSpeechBubble.offsetHeight - 15) + "px";
                }, 0);
            }
        }
    });
}

/* ===================== DOWNLOAD CHART ===================== */
downloadBtn.addEventListener("click", () => {
    const auditName = auditSelect.value || "chart";
    const link = document.createElement("a");
    link.download = `score-audit-${auditName.toLowerCase()}.png`;
    link.href = canvas.toDataURL("image/png");
    link.click();
});

/* ===================== LOGOUT ===================== */
logoutBtn.addEventListener("click", async () => {
    sessionStorage.clear();
    clearAuditProgress();
    await supabase.auth.signOut();
    window.location.href = "login.html";
});
