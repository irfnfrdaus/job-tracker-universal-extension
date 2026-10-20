const BRIDGE_URL = "http://127.0.0.1:5252";

let isBridgeConnected = false;

document.addEventListener("DOMContentLoaded", () => {
  setupTabs();
  initDefaults();
  setupEventListeners();

  // Async tasks in background
  checkBridgeStatus();
  autoDetectJobDetails();
  loadRecentApplications();
});

// Setup Tab Switching
function setupTabs() {
  const tabBtns = document.querySelectorAll(".tab-btn");
  tabBtns.forEach(btn => {
    btn.addEventListener("click", () => {
      tabBtns.forEach(b => b.classList.remove("active"));
      document.querySelectorAll(".tab-content").forEach(c => c.classList.remove("active"));
      
      btn.classList.add("active");
      const targetId = btn.getAttribute("data-tab");
      document.getElementById(targetId).classList.add("active");
    });
  });
}

// Initialize Default Values
function initDefaults() {
  const today = new Date().toISOString().split("T")[0];
  document.getElementById("dateApplied").value = today;
}

// Check Python Bridge Connectivity (Optional for local Excel users)
async function checkBridgeStatus() {
  const pill = document.getElementById("statusPill");
  const text = document.getElementById("statusText");
  const bridgeDetail = document.getElementById("bridgeDetail");

  try {
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), 1200);

    const response = await fetch(`${BRIDGE_URL}/status`, { signal: controller.signal });
    clearTimeout(timeoutId);

    if (response.ok) {
      const data = await response.json();
      isBridgeConnected = true;
      pill.className = "status-pill";
      text.innerText = "Excel Connected";
      if (bridgeDetail) {
        bridgeDetail.innerHTML = `<span style="color: #10B981; font-weight: 600;">Active</span> • File: <code>${data.excel_path}</code><br>Total Records: <b>${data.total_records}</b>`;
      }
      syncOfflineJobsToExcel();
      return true;
    }
  } catch (err) {
    // Offline - perfectly fine, extension operates standalone
  }

  isBridgeConnected = false;
  pill.className = "status-pill offline";
  text.innerText = "Browser Local";
  if (bridgeDetail) {
    bridgeDetail.innerHTML = `<span style="color: #64748B; font-weight: 600;">Standalone Mode</span>. Applications save directly inside your browser storage.<br>Run <code>start_bridge</code> to link with desktop Excel.`;
  }
  return false;
}

// Auto-sync any jobs saved while offline to desktop Excel
async function syncOfflineJobsToExcel() {
  if (!isBridgeConnected) return;
  chrome.storage.local.get({ jobs: [] }, async (res) => {
    const jobs = res.jobs || [];
    let syncedAny = false;

    for (let i = 0; i < jobs.length; i++) {
      const j = jobs[i];
      if (j.app_id && j.app_id.startsWith("LOCAL-")) {
        try {
          const response = await fetch(`${BRIDGE_URL}/add-job`, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(j)
          });
          if (response.ok) {
            const result = await response.json();
            jobs[i].app_id = result.app_id;
            syncedAny = true;
          }
        } catch (err) {
          console.warn("Could not sync offline job:", err);
        }
      }
    }

    if (syncedAny) {
      chrome.storage.local.set({ jobs }, () => {
        loadRecentApplications();
        showToast("Synced offline applications to Excel!", "success");
      });
    }
  });
}

// Auto-Detect Information from Active Tab
async function autoDetectJobDetails() {
  try {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    if (!tab || !tab.url) return;

    document.getElementById("url").value = tab.url;

    // Detect Platform
    const urlLower = tab.url.toLowerCase();
    const platformSelect = document.getElementById("platform");
    if (urlLower.includes("linkedin.com")) {
      platformSelect.value = "LinkedIn";
    } else if (urlLower.includes("jobstreet.com")) {
      platformSelect.value = "JobStreet";
    } else if (urlLower.includes("indeed.com")) {
      platformSelect.value = "Indeed";
    } else if (urlLower.includes("glassdoor.com")) {
      platformSelect.value = "Glassdoor";
    } else if (urlLower.includes("myfuturejobs")) {
      platformSelect.value = "MyFutureJobs";
    } else if (urlLower.includes("handshake")) {
      platformSelect.value = "Handshake";
    } else {
      platformSelect.value = "Company Portal";
    }

    // Attempt DOM Scrape via Scripting API
    if (tab.id && !tab.url.startsWith("chrome://") && !tab.url.startsWith("edge://") && !tab.url.startsWith("opera://") && !tab.url.startsWith("about:")) {
      chrome.scripting.executeScript({
        target: { tabId: tab.id },
        func: scrapePageData
      }, (results) => {
        if (chrome.runtime.lastError || !results || !results[0] || !results[0].result) {
          parseFallbackFromTitle(tab.title);
          return;
        }

        const data = results[0].result;
        if (data.company) document.getElementById("company").value = data.company;
        if (data.title) document.getElementById("jobTitle").value = data.title;
        if (data.location) document.getElementById("location").value = data.location;
        if (data.salary) document.getElementById("salary").value = data.salary;

        if (!data.company || !data.title) {
          parseFallbackFromTitle(tab.title);
        }
      });
    } else {
      parseFallbackFromTitle(tab.title);
    }
  } catch (err) {
    console.warn("Auto-detect failed:", err);
  }
}

// Function injected into the target page
function scrapePageData() {
  const data = { company: "", title: "", location: "", salary: "" };
  const url = window.location.href.toLowerCase();

  if (url.includes("linkedin.com")) {
    const t = document.querySelector(".job-details-jobs-unified-top-card__job-title, .jobs-unified-top-card__job-title, h1.t-24");
    if (t) data.title = t.innerText.trim();

    const c = document.querySelector(".job-details-jobs-unified-top-card__company-name, .jobs-unified-top-card__company-name, .job-details-jobs-unified-top-card__primary-description a");
    if (c) data.company = c.innerText.trim();

    const loc = document.querySelector(".job-details-jobs-unified-top-card__primary-description span:nth-of-type(1), .jobs-unified-top-card__bullet");
    if (loc) data.location = loc.innerText.trim();
  } 
  else if (url.includes("jobstreet.com")) {
    const t = document.querySelector('[data-automation="job-detail-title"], h1');
    if (t) data.title = t.innerText.trim();

    const c = document.querySelector('[data-automation="advertiser-name"]');
    if (c) data.company = c.innerText.trim();

    const loc = document.querySelector('[data-automation="job-detail-location"]');
    if (loc) data.location = loc.innerText.trim();
  }
  else if (url.includes("indeed.com")) {
    const t = document.querySelector(".jobsearch-JobInfoHeader-title, h1");
    if (t) data.title = t.innerText.trim();

    const c = document.querySelector('[data-company-name="true"], .jobsearch-CompanyInfoContainer a');
    if (c) data.company = c.innerText.trim();

    const loc = document.querySelector(".jobsearch-JobInfoHeader-subtitle > div:last-child");
    if (loc) data.location = loc.innerText.trim();
  }

  return data;
}

// Fallback parsing from Tab Title (e.g., "PETRONAS hiring Data Analyst in Kuala Lumpur | LinkedIn")
function parseFallbackFromTitle(rawTitle) {
  if (!rawTitle) return;
  const title = rawTitle.replace(/\s*[-–|]\s*(LinkedIn|JobStreet|Indeed|Glassdoor).*$/i, "").trim();

  // Pattern: "Job Title at Company" or "Company hiring Job Title"
  if (title.includes(" at ")) {
    const parts = title.split(" at ");
    if (!document.getElementById("jobTitle").value) document.getElementById("jobTitle").value = parts[0].trim();
    if (!document.getElementById("company").value) document.getElementById("company").value = parts[1].trim();
  } else if (title.includes(" hiring ")) {
    const parts = title.split(" hiring ");
    if (!document.getElementById("company").value) document.getElementById("company").value = parts[0].trim();
    if (!document.getElementById("jobTitle").value) document.getElementById("jobTitle").value = parts[1].trim();
  } else if (!document.getElementById("jobTitle").value) {
    document.getElementById("jobTitle").value = title;
  }
}

// Setup Event Listeners
function setupEventListeners() {
  const statusPill = document.getElementById("statusPill");
  if (statusPill) statusPill.addEventListener("click", checkBridgeStatus);

  const retryBtn = document.getElementById("retryBridgeBtn");
  if (retryBtn) retryBtn.addEventListener("click", checkBridgeStatus);

  const jobForm = document.getElementById("jobForm");
  if (jobForm) jobForm.addEventListener("submit", handleSubmit);

  const openDashboardBtn = document.getElementById("openDashboardBtn");
  if (openDashboardBtn) {
    openDashboardBtn.addEventListener("click", (e) => {
      e.preventDefault();
      openDashboard();
    });
  }

  const openExcelBtn = document.getElementById("openExcelBtn");
  if (openExcelBtn) {
    openExcelBtn.addEventListener("click", (e) => {
      e.preventDefault();
      openExcelOrDashboard();
    });
  }

  const copyBtn = document.getElementById("copyDetailsBtn");
  if (copyBtn) copyBtn.addEventListener("click", copyApplicationDetails);

  const exportBtn = document.getElementById("exportCsvBtn");
  if (exportBtn) exportBtn.addEventListener("click", exportCsv);
}

// Open In-Browser Dashboard (Works standalone on Mac/Windows/Linux)
function openDashboard() {
  const dashboardUrl = chrome.runtime.getURL("dashboard.html");
  chrome.tabs.create({ url: dashboardUrl });
}

// Open Excel directly if bridge is connected, else open Dashboard
async function openExcelOrDashboard() {
  if (isBridgeConnected) {
    try {
      showToast("Triggering Excel on desktop...", "success");
      await fetch(`${BRIDGE_URL}/open-excel`);
      chrome.tabs.create({ url: `${BRIDGE_URL}/view` });
      return;
    } catch (err) {
      console.warn("Failed to reach bridge:", err);
    }
  }

  // Fallback to standalone dashboard
  openDashboard();
  showToast("Opened standalone Dashboard! Run bridge to sync with Excel.", "success");
}

// Handle Form Submission
async function handleSubmit(e) {
  e.preventDefault();

  const job = {
    company: document.getElementById("company").value.trim(),
    job_title: document.getElementById("jobTitle").value.trim(),
    platform: document.getElementById("platform").value,
    status: document.getElementById("status").value,
    job_type: document.getElementById("jobType").value,
    location: document.getElementById("location").value.trim() || "Remote / Hybrid",
    date_applied: document.getElementById("dateApplied").value,
    salary: document.getElementById("salary").value.trim(),
    email: document.getElementById("email").value.trim(),
    url: document.getElementById("url").value.trim(),
    notes: document.getElementById("notes").value.trim(),
    created_at: new Date().toISOString()
  };

  const saveBtn = document.getElementById("saveBtn");
  saveBtn.disabled = true;
  saveBtn.innerText = "Saving...";

  // 1. If Bridge is Connected, send to desktop Excel
  if (isBridgeConnected) {
    try {
      const response = await fetch(`${BRIDGE_URL}/add-job`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(job)
      });

      if (response.ok) {
        const result = await response.json();
        job.app_id = result.app_id;
        await saveToLocalStorage(job);
        await loadRecentApplications();
        showToast(result.message || `Saved to Excel as ${result.app_id}!`, "success");
        resetFormExceptDates();
        saveBtn.disabled = false;
        saveBtn.innerText = "Save Application";
        return;
      }
    } catch (err) {
      console.warn("Bridge save failed, falling back to local storage:", err);
    }
  }

  // 2. Standalone Save to chrome.storage.local
  const total = await getLocalCount();
  job.app_id = `APP-${String(total + 1).padStart(3, "0")}`;
  await saveToLocalStorage(job);
  await loadRecentApplications();
  showToast(`Application saved as ${job.app_id}!`, "success");
  resetFormExceptDates();
  saveBtn.disabled = false;
  saveBtn.innerText = "Save Application";
}

// Reset Form fields
function resetFormExceptDates() {
  document.getElementById("company").value = "";
  document.getElementById("jobTitle").value = "";
  document.getElementById("salary").value = "";
  document.getElementById("email").value = "";
  document.getElementById("notes").value = "";
}

// Helper: Get Count of Stored Jobs
async function getLocalCount() {
  return new Promise((resolve) => {
    chrome.storage.local.get({ jobs: [] }, (res) => {
      resolve(res.jobs.length);
    });
  });
}

// Save job record to browser local storage
async function saveToLocalStorage(job) {
  return new Promise((resolve) => {
    chrome.storage.local.get({ jobs: [] }, (res) => {
      const jobs = res.jobs;
      // Prepend so newest is first
      jobs.unshift(job);
      chrome.storage.local.set({ jobs }, () => resolve());
    });
  });
}

// Load Recent Applications into Tab 2
async function loadRecentApplications() {
  chrome.storage.local.get({ jobs: [] }, (res) => {
    const jobs = res.jobs;
    const countEl = document.getElementById("recentCount");
    const container = document.getElementById("recentContainer");

    if (countEl) countEl.innerText = jobs.length;
    if (!container) return;

    if (jobs.length === 0) {
      container.innerHTML = `<p style="text-align: center; color: var(--text-muted); padding: 24px;">No applications recorded yet.<br>Apply on any job portal and click Save!</p>`;
      return;
    }

    let html = "";
    jobs.slice(0, 10).forEach(j => {
      let badgeStyle = "background: #E2E8F0; color: #475569;";
      if (j.status === "Applied") badgeStyle = "background: #DBEAFE; color: #1E40AF;";
      else if (j.status === "Interviewing") badgeStyle = "background: #FEF3C7; color: #92400E;";
      else if (j.status === "Offer Received") badgeStyle = "background: #D1FAE5; color: #065F46;";
      else if (j.status === "Rejected") badgeStyle = "background: #FEE2E2; color: #991B1B;";

      html += `
        <div class="recent-item">
          <div class="recent-item-header">
            <span class="recent-company">${escapeHtml(j.company)}</span>
            <span class="recent-status" style="${badgeStyle}">${escapeHtml(j.status)}</span>
          </div>
          <div class="recent-title">${escapeHtml(j.job_title)}</div>
          <div class="recent-meta">
            <span>📅 ${j.date_applied || "N/A"}</span>
            <span>🌐 ${escapeHtml(j.platform)}</span>
            ${j.salary ? `<span>💰 ${escapeHtml(j.salary)}</span>` : ""}
          </div>
        </div>
      `;
    });

    container.innerHTML = html;
  });
}

// Export All Stored Jobs to CSV File
function exportCsv() {
  chrome.storage.local.get({ jobs: [] }, (res) => {
    const jobs = res.jobs;
    if (jobs.length === 0) {
      showToast("No applications to export!", "error");
      return;
    }

    const headers = ["App ID", "Date Applied", "Company", "Job Title", "Job Type", "Location", "Platform", "Status", "Job URL", "Email", "Salary", "Notes"];
    const csvRows = [headers.join(",")];

    jobs.forEach(j => {
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
    a.download = `Job_Applications_${new Date().toISOString().split("T")[0]}.csv`;
    a.click();
    URL.revokeObjectURL(url);
    showToast("Downloaded CSV backup successfully!", "success");
  });
}

// Copy Application Details to Clipboard
function copyApplicationDetails() {
  const company = document.getElementById("company").value.trim();
  const title = document.getElementById("jobTitle").value.trim();
  const platform = document.getElementById("platform").value;
  const status = document.getElementById("status").value;
  const date = document.getElementById("dateApplied").value;
  const url = document.getElementById("url").value.trim();

  if (!company || !title) {
    showToast("Please fill Company and Title first!", "error");
    return;
  }

  const text = `📌 Job Application: ${title} at ${company}\nPlatform: ${platform}\nStatus: ${status}\nDate: ${date}\nLink: ${url || "N/A"}`;
  navigator.clipboard.writeText(text).then(() => {
    showToast("Copied application summary to clipboard!", "success");
  }).catch(() => {
    showToast("Failed to copy to clipboard", "error");
  });
}

// Toast helper
function showToast(message, type = "success") {
  const toast = document.getElementById("toast");
  toast.innerText = message;
  toast.className = type;
  setTimeout(() => {
    toast.className = "";
  }, 4000);
}

// Escape HTML helper
function escapeHtml(str) {
  if (!str) return "";
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}
