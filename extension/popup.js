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

// Check Python Bridge Connectivity
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
      bridgeDetail.innerHTML = `<span style="color: #10B981; font-weight: 600;">Connected</span> to <code>${data.excel_path}</code><br>Total Records: <b>${data.total_records}</b>`;
      // Auto-sync any jobs saved while offline
      syncOfflineJobsToExcel();
      return true;
    }
  } catch (err) {
    // Offline
  }

  isBridgeConnected = false;
  pill.className = "status-pill offline";
  text.innerText = "Offline (Local)";
  bridgeDetail.innerHTML = `<span style="color: #F59E0B; font-weight: 600;">Offline</span> (Bridge not running).<br>Run <code>Start_Job_Tracker.bat</code> to connect live with Excel.`;
  return false;
}

// Auto-sync any offline/local jobs into Excel when bridge connects
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
    } else if (urlLower.includes("myfuturejobs")) {
      platformSelect.value = "MyFutureJobs";
    } else if (urlLower.includes("glassdoor.com")) {
      platformSelect.value = "Glassdoor";
    } else {
      platformSelect.value = "Company Portal";
    }

    // Attempt DOM Scrape via Scripting API
    if (tab.id && !tab.url.startsWith("chrome://") && !tab.url.startsWith("edge://") && !tab.url.startsWith("opera://")) {
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

  const openTrackerBtn = document.getElementById("openTrackerBtn");
  if (openTrackerBtn) {
    openTrackerBtn.addEventListener("click", (e) => {
      e.preventDefault();
      e.stopPropagation();
      openTrackerDashboard(false);
    });
  }

  const openExcelDirectBtn = document.getElementById("openExcelDirectBtn");
  if (openExcelDirectBtn) {
    openExcelDirectBtn.addEventListener("click", (e) => {
      e.preventDefault();
      e.stopPropagation();
      openTrackerDashboard(true);
    });
  }

  const openExcelBtn = document.getElementById("openExcelBtn");
  if (openExcelBtn) {
    openExcelBtn.addEventListener("click", (e) => {
      e.preventDefault();
      e.stopPropagation();
      openTrackerDashboard(true);
    });
  }

  const copyBtn = document.getElementById("copyDetailsBtn");
  if (copyBtn) copyBtn.addEventListener("click", copyApplicationDetails);

  const exportBtn = document.getElementById("exportCsvBtn");
  if (exportBtn) exportBtn.addEventListener("click", exportCsv);
}

// Open Live Tracker Dashboard Tab & Excel File
async function openTrackerDashboard(launchExcel = true) {
  try {
    // 1. Immediately open the rich live tracker dashboard in a new Opera tab
    if (chrome && chrome.tabs && chrome.tabs.create) {
      chrome.tabs.create({ url: `${BRIDGE_URL}/view` });
    } else {
      window.open(`${BRIDGE_URL}/view`, "_blank");
    }

    // 2. Request Bridge to trigger Excel or Explorer
    if (launchExcel) {
      showToast("Opening Tracker Dashboard & Excel...", "success");
      try {
        const res = await fetch(`${BRIDGE_URL}/open-excel`);
        const data = await res.json();
        if (data.success) {
          isBridgeConnected = true;
          const pill = document.getElementById("statusPill");
          if (pill) {
            pill.className = "status-pill";
            document.getElementById("statusText").innerText = "Excel Connected";
          }
        }
      } catch (err) {
        console.warn("Could not trigger Excel via bridge:", err);
      }
    } else {
      showToast("Opened Live Tracker Tab!", "success");
    }
  } catch (err) {
    showToast("Bridge offline! Run Start_Job_Tracker.bat on C: drive", "error");
  }
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
    location: document.getElementById("location").value.trim() || "Kuala Lumpur, Malaysia",
    date_applied: document.getElementById("dateApplied").value,
    salary: document.getElementById("salary").value.trim(),
    email: document.getElementById("email").value.trim(),
    url: document.getElementById("url").value.trim(),
    notes: document.getElementById("notes").value.trim()
  };

  const saveBtn = document.getElementById("saveBtn");
  saveBtn.disabled = true;
  saveBtn.innerText = "Saving...";

  // 1. Try sending to Python Bridge
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
        showToast(result.message || `Saved to Tracker as ${result.app_id}!`, "success");
        await saveToLocalStorage(job);
        await loadRecentApplications();
        resetFormExceptDates();
        saveBtn.disabled = false;
        saveBtn.innerText = "Save to Excel Tracker";
        return;
      }
    } catch (err) {
      console.warn("Failed to reach bridge, falling back to local storage:", err);
    }
  }

  // 2. Offline Fallback to chrome.storage.local
  job.app_id = `LOCAL-${Date.now().toString().slice(-4)}`;
  await saveToLocalStorage(job);
  await loadRecentApplications();
  showToast(`Saved locally as ${job.app_id}! (Will sync to Excel when bridge runs)`, "success");
  resetFormExceptDates();
  saveBtn.disabled = false;
  saveBtn.innerText = "Save to Excel Tracker";
}

// Reset Form fields
function resetFormExceptDates() {
  document.getElementById("company").value = "";
  document.getElementById("jobTitle").value = "";
  document.getElementById("salary").value = "";
  document.getElementById("email").value = "";
  document.getElementById("notes").value = "";
}

// Toast helper
function showToast(message, type = "success") {
  const toast = document.getElementById("toast");
  toast.innerText = message;
  toast.className = type;
  setTimeout(() => {
    toast.className = "";
  }, 4500);
}

// Save in Chrome/Opera local storage
async function saveToLocalStorage(job) {
  return new Promise((resolve) => {
    chrome.storage.local.get(["applications"], (result) => {
      const list = result.applications || [];
      list.unshift(job);
      chrome.storage.local.set({ applications: list }, () => resolve());
    });
  });
}

// Load Recent Applications into the UI
async function loadRecentApplications() {
  chrome.storage.local.get(["applications"], (result) => {
    const list = result.applications || [];
    document.getElementById("recentCount").innerText = list.length;
    const container = document.getElementById("recentContainer");

    if (list.length === 0) {
      container.innerHTML = '<p style="text-align: center; color: var(--text-muted); padding: 20px;">No applications recorded yet.</p>';
      return;
    }

    container.innerHTML = list.slice(0, 15).map(item => {
      const sClean = (item.status || "default").replace(/ \/ /g, "---").replace(/ /g, "-");
      const statusClass = `status-badge-${sClean}`;
      return `
      <div class="recent-item">
        <div class="recent-item-header">
          <span class="recent-company">${escapeHtml(item.company)}</span>
          <span class="recent-status ${statusClass}">${escapeHtml(item.status || "Applied")}</span>
        </div>
        <div class="recent-title">${escapeHtml(item.job_title)}</div>
        <div class="recent-meta">
          <span>${item.date_applied || ""}</span>
          <span>•</span>
          <span>${escapeHtml(item.platform || "")}</span>
          <span>•</span>
          <span>${item.app_id || ""}</span>
        </div>
      </div>
    `;
    }).join("");
  });
}

// Copy Summary to Clipboard
function copyApplicationDetails() {
  const comp = document.getElementById("company").value.trim() || "Company";
  const title = document.getElementById("jobTitle").value.trim() || "Role";
  const url = document.getElementById("url").value.trim();
  const date = document.getElementById("dateApplied").value;
  const status = document.getElementById("status").value;

  const text = `Application Log: ${comp} | ${title}\nDate: ${date} | Status: ${status}\nURL: ${url}`;
  navigator.clipboard.writeText(text).then(() => {
    showToast("Application summary copied to clipboard!", "success");
  });
}

// Export CSV Backup
function exportCsv() {
  chrome.storage.local.get(["applications"], (result) => {
    const list = result.applications || [];
    if (list.length === 0) {
      alert("No applications saved locally yet.");
      return;
    }

    const headers = ["App ID", "Date", "Company", "Job Title", "Platform", "Status", "Salary", "Location", "URL", "Notes"];
    const rows = list.map(j => [
      j.app_id || "",
      j.date_applied || "",
      `"${(j.company || '').replace(/"/g, '""')}"`,
      `"${(j.job_title || '').replace(/"/g, '""')}"`,
      j.platform || "",
      j.status || "",
      j.salary || "",
      `"${(j.location || '').replace(/"/g, '""')}"`,
      `"${(j.url || '').replace(/"/g, '""')}"`,
      `"${(j.notes || '').replace(/"/g, '""')}"`
    ]);

    const csvContent = "data:text/csv;charset=utf-8," + [headers.join(","), ...rows.map(r => r.join(","))].join("\n");
    const encodedUri = encodeURI(csvContent);
    const link = document.createElement("a");
    link.setAttribute("href", encodedUri);
    link.setAttribute("download", `Job_Applications_Backup_${new Date().toISOString().slice(0, 10)}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  });
}

function escapeHtml(text) {
  if (!text) return "";
  const div = document.createElement("div");
  div.innerText = text;
  return div.innerHTML;
}
