(() => {
  "use strict";

  const AUTO_REFRESH_MS = 15 * 60 * 1000;
  const LOAD_WARNING_MS = 45 * 1000;
  const header = document.getElementById("app-header");
  const shell = document.getElementById("report-shell");
  const frame = document.getElementById("report-frame");
  const refreshButton = document.getElementById("refresh-report");
  const retryButton = document.getElementById("retry-report");
  const fullscreenButton = document.getElementById("fullscreen-report");
  const fullscreenLabel = document.getElementById("fullscreen-label");
  const openPowerBi = document.getElementById("open-powerbi");
  const statusDot = document.getElementById("status-dot");
  const connectionStatus = document.getElementById("connection-status");
  const refreshStatus = document.getElementById("refresh-status");
  const loadingTitle = document.getElementById("loading-title");
  const loadingMessage = document.getElementById("loading-message");
  const reportUrl = frame.dataset.reportUrl;

  let nextRefreshAt = Date.now() + AUTO_REFRESH_MS;
  let loadWarningTimer = null;
  let refreshing = false;
  let wasOffline = !navigator.onLine;

  openPowerBi.href = reportUrl;

  function setHeaderHeight() {
    document.documentElement.style.setProperty("--header-height", `${Math.ceil(header.getBoundingClientRect().height)}px`);
  }

  function dhakaTime(date = new Date()) {
    return new Intl.DateTimeFormat("en-GB", {
      hour: "2-digit",
      minute: "2-digit",
      timeZone: "Asia/Dhaka",
    }).format(date);
  }

  function countdown(milliseconds) {
    const seconds = Math.max(0, Math.ceil(milliseconds / 1000));
    const minutes = Math.floor(seconds / 60);
    return `${minutes}:${String(seconds % 60).padStart(2, "0")}`;
  }

  function setLoadingState() {
    shell.classList.remove("loaded");
    statusDot.className = "status-dot loading";
    connectionStatus.textContent = navigator.onLine ? "Refreshing Power BI" : "Internet unavailable";
    loadingTitle.textContent = navigator.onLine ? "Opening Receiving Dashboard" : "Waiting for internet connection";
    loadingMessage.textContent = navigator.onLine
      ? "Connecting to the latest published Power BI report…"
      : "The report will reconnect automatically when this device is online.";
    retryButton.hidden = true;
  }

  function startLoadWarning() {
    clearTimeout(loadWarningTimer);
    loadWarningTimer = setTimeout(() => {
      if (shell.classList.contains("loaded")) return;
      connectionStatus.textContent = "Report is taking longer";
      loadingTitle.textContent = "Power BI is still loading";
      loadingMessage.textContent = "Check the internet connection, try again, or open the original Power BI report.";
      retryButton.hidden = false;
    }, LOAD_WARNING_MS);
  }

  function reloadReport(source = "manual") {
    if (refreshing || !navigator.onLine) {
      if (!navigator.onLine) setLoadingState();
      return;
    }

    refreshing = true;
    setLoadingState();
    refreshButton.disabled = true;
    refreshButton.classList.add("refreshing");
    nextRefreshAt = Date.now() + AUTO_REFRESH_MS;
    startLoadWarning();

    frame.src = reportUrl;

    window.setTimeout(() => {
      refreshing = false;
      refreshButton.disabled = false;
      refreshButton.classList.remove("refreshing");
      if (source === "automatic") refreshStatus.textContent = "Scheduled refresh requested";
    }, 1200);
  }

  function updateOnlineState() {
    if (!navigator.onLine) {
      wasOffline = true;
      statusDot.className = "status-dot offline";
      connectionStatus.textContent = "Offline";
      refreshStatus.textContent = "Waiting to reconnect";
      return;
    }

    if (wasOffline) {
      wasOffline = false;
      reloadReport("reconnect");
    }
  }

  function tick() {
    updateOnlineState();
    if (!navigator.onLine) return;

    const remaining = nextRefreshAt - Date.now();
    if (remaining <= 0) {
      reloadReport("automatic");
      return;
    }

    if (shell.classList.contains("loaded")) {
      refreshStatus.textContent = `Next refresh in ${countdown(remaining)}`;
    }
  }

  frame.addEventListener("load", () => {
    clearTimeout(loadWarningTimer);
    window.setTimeout(() => {
      shell.classList.add("loaded");
      statusDot.className = "status-dot";
      connectionStatus.textContent = "Power BI live";
      refreshStatus.textContent = `Checked ${dhakaTime()} · refresh in ${countdown(nextRefreshAt - Date.now())}`;
      retryButton.hidden = true;
    }, 350);
  });

  refreshButton.addEventListener("click", () => reloadReport("manual"));
  retryButton.addEventListener("click", () => reloadReport("retry"));

  fullscreenButton.addEventListener("click", async () => {
    try {
      if (document.fullscreenElement) await document.exitFullscreen();
      else await document.documentElement.requestFullscreen();
    } catch {
      window.open(reportUrl, "_blank", "noopener,noreferrer");
    }
  });

  document.addEventListener("fullscreenchange", () => {
    const active = Boolean(document.fullscreenElement);
    fullscreenLabel.textContent = active ? "Exit full screen" : "Full screen";
    fullscreenButton.title = active ? "Exit full-screen view" : "Open the dashboard in full screen";
    setHeaderHeight();
  });

  document.addEventListener("visibilitychange", () => {
    if (!document.hidden && Date.now() >= nextRefreshAt) reloadReport("automatic");
  });

  window.addEventListener("online", updateOnlineState);
  window.addEventListener("offline", updateOnlineState);
  window.addEventListener("resize", setHeaderHeight);

  if ("ResizeObserver" in window) new ResizeObserver(setHeaderHeight).observe(header);
  setHeaderHeight();
  startLoadWarning();
  window.setInterval(tick, 1000);
})();
