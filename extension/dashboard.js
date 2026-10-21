let allJobs = [];

document.addEventListener("DOMContentLoaded", () => {
  loadJobs();
  setupEventListeners();
});

function setupEventListeners() {
  document.getElementById("searchInput").addEventListener("input", renderTable);
  document.getElementById("statusFilter").addEventListener("change", renderTable);
  document.getElementById("exportCsvBtn").addEventListener("click", exportCsv);
  document.getElementById("refreshBtn").addEventListener("click", loadJobs);
}

function loadJobs() {
  chrome.storage.local.get({ jobs: [] }, (res) => {
    allJobs = res.jobs || [];
    updateStats();
    renderTable();
  });
}

function updateStats() {
  const total = allJobs.length;
  let applied = 0, interview = 0, offers = 0, rejected = 0;

  allJobs.forEach(j => {
    const s = (j.status || "").toLowerCase();
    if (s.includes("applied") || s.includes("review")) applied++;
    else if (s.includes("interview") || s.includes("test") || s.includes("assessment")) interview++;
    else if (s.includes("offer") || s.includes("shortlist")) offers++;
    else if (s.includes("reject")) rejected++;
  });

  document.getElementById("statTotal").innerText = total;
  document.getElementById("statApplied").innerText = applied;
  document.getElementById("statInterview").innerText = interview;
  document.getElementById("statOffers").innerText = offers;
  document.getElementById("statRejected").innerText = rejected;
}

function renderTable() {
  const query = document.getElementById("searchInput").value.toLowerCase().trim();
  const filter = document.getElementById("statusFilter").value;
  const tbody = document.getElementById("tableBody");

  const filtered = allJobs.filter(j => {
    const matchesFilter = (filter === "ALL") || (j.status && j.status.includes(filter));
    const textCorpus = `${j.app_id || ""} ${j.company || ""} ${j.job_title || ""} ${j.platform || ""} ${j.location || ""} ${j.notes || ""}`.toLowerCase();
    const matchesQuery = !query || textCorpus.includes(query);
    return matchesFilter && matchesQuery;
  });

  if (filtered.length === 0) {
    tbody.innerHTML = `<tr><td colspan="12" style="text-align: center; padding: 40px; color: var(--text-muted);">No applications found matching your search.</td></tr>`;
    return;
  }

  let html = "";
  filtered.forEach(j => {
    const statusText = j.status || "Applied";
    let statusClass = "status-Applied";
    if (statusText.includes("Interview") || statusText.includes("Test")) statusClass = "status-Interviewing";
    else if (statusText.includes("Offer")) statusClass = "status-Offer-Received";
    else if (statusText.includes("Reject")) statusClass = "status-Rejected";
    else if (statusText.includes("Wishlist") || statusText.includes("Target")) statusClass = "status-Wishlist";

    const STATUS_OPTIONS = [
      "Applied",
      "Under Review",
      "Assessment / Test",
      "Interviewing",
      "Shortlisted",
      "Offer Received",
      "Target / Backlog",
      "Rejected"
    ];

    const opts = STATUS_OPTIONS.map(s => `<option value="${s}" ${s === statusText ? 'selected' : ''}>${s}</option>`).join("");

    html += `
      <tr>
        <td class="mono" style="color: #60A5FA; font-weight: bold;">${escapeHtml(j.app_id || "APP-000")}</td>
        <td class="mono" style="color: var(--text-muted);">${escapeHtml(j.date_applied || "N/A")}</td>
        <td style="color: #FFFFFF; font-weight: 700;">${escapeHtml(j.company)}</td>
        <td style="color: #E2E8F0; font-weight: 500;">${escapeHtml(j.job_title)}</td>
        <td style="color: var(--text-muted);">${escapeHtml(j.job_type || "-")}</td>
        <td style="color: var(--text-muted);">${escapeHtml(j.location || "-")}</td>
        <td><span style="color: #CBD5E1;">${escapeHtml(j.platform || "-")}</span></td>
        <td>
          <select class="status-select" onchange="changeJobStatus('${escapeHtml(j.app_id)}', this.value)">
            ${opts}
          </select>
        </td>
        <td>${j.url ? `<a href="${escapeHtml(j.url)}" target="_blank" class="link-btn">Link ↗</a>` : `<span style="color: #475569;">-</span>`}</td>
        <td class="mono" style="color: #34D399; font-weight: 600;">${escapeHtml(j.salary || "-")}</td>
        <td style="color: var(--text-muted); max-width: 280px; font-size: 11px;">${escapeHtml(j.notes || "-")}</td>
        <td style="text-align: center;">
          <button class="btn-delete" title="Delete application" onclick="deleteJob('${escapeHtml(j.app_id)}')">🗑️</button>
        </td>
      </tr>
    `;
  });

  tbody.innerHTML = html;
}

window.changeJobStatus = async function(appId, newStatus) {
  const job = allJobs.find(j => j.app_id === appId);
  if (job) {
    job.status = newStatus;
    chrome.storage.local.set({ jobs: allJobs }, () => {
      updateStats();
      renderTable();
    });

    // If bridge is running, sync status update to desktop Excel too!
    try {
      await fetch("http://127.0.0.1:5252/update-status", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ app_id: appId, status: newStatus })
      });
    } catch (e) {
      // Bridge is optional, ignore if bridge is offline
    }
  }
};

window.deleteJob = function(appId) {
  if (!confirm(`Are you sure you want to delete application ${appId}?`)) return;
  allJobs = allJobs.filter(j => j.app_id !== appId);
  chrome.storage.local.set({ jobs: allJobs }, () => {
    updateStats();
    renderTable();
  });
};

function exportCsv() {
  if (allJobs.length === 0) {
    alert("No applications to export!");
    return;
  }

  const headers = ["App ID", "Date Applied", "Company", "Job Title", "Job Type", "Location", "Platform", "Status", "Job URL", "Email", "Salary", "Notes"];
  const csvRows = [headers.join(",")];

  allJobs.forEach(j => {
    const row = [
      `"${(j.app_id || "").replace(/"/g, '""')}"`,
      `"${(j.date_applied || "").replace(/"/g, '""')}"`,
      `"${(j.company || "").replace(/"/g, '""')}"`,
      `"${(j.job_title || "").replace(/"/g, '""')}"`,
      `"${(j.job_type || "").replace(/"/g, '""')}"`,
      `"${(j.location || "").replace(/"/g, '""')}"`,
      `"${(j.platform || "").replace(/"/g, '""')}"`,
      `"${(j.status || "").replace(/"/g, '""')}"`,
      `"${(j.url || "").replace(/"/g, '""')}"`,
      `"${(j.email || "").replace(/"/g, '""')}"`,
      `"${(j.salary || "").replace(/"/g, '""')}"`,
      `"${(j.notes || "").replace(/"/g, '""')}"`
    ];
    csvRows.push(row.join(","));
  });

  const blob = new Blob([csvRows.join("\n")], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `Job_Applications_Export_${new Date().toISOString().split("T")[0]}.csv`;
  a.click();
  URL.revokeObjectURL(url);
}

function escapeHtml(str) {
  if (!str) return "";
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}
