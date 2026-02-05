// SP Reveal – popup script
// Behavior required by Rik:
// - On popup open: show "SP Reveal — Productivity tool for SharePoint"
// - After any button click: replace with action status (like before)
// - On next open: reset to branding text again

(() => {
  const BRANDING = "SP Reveal — Productivity tool for SharePoint";

  // --- DOM helpers ---
  function $(sel) {
    return document.querySelector(sel);
  }
  function $all(sel) {
    return Array.from(document.querySelectorAll(sel));
  }

  function setStatus(msg) {
    const el = $("#status");
    if (!el) return;
    el.textContent = msg || BRANDING;
  }

  // --- Chrome tab + messaging helpers ---
  async function activeTabId() {
    const [tab] = await chrome.tabs.query({
      active: true,
      currentWindow: true,
    });
    return tab?.id;
  }

  async function sendCmd(cmd, args = {}) {
    const tabId = await activeTabId();
    if (!tabId) throw new Error("No active tab to send command to.");
    return new Promise((resolve, reject) => {
      chrome.tabs.sendMessage(tabId, { cmd, args }, (res) => {
        const err = chrome.runtime.lastError;
        if (err) return reject(new Error(err.message || "sendMessage failed"));
        resolve(res);
      });
    });
  }

  // --- Clipboard helper ---
  async function copyToClipboard(text) {
    try {
      if (navigator.clipboard && window.isSecureContext) {
        await navigator.clipboard.writeText(String(text));
        setStatus("Copied");
        return true;
      }
    } catch {
      // fall through to legacy path
    }
    try {
      const ta = document.createElement("textarea");
      ta.style.position = "fixed";
      ta.style.top = "-9999px";
      ta.value = String(text);
      document.body.appendChild(ta);
      ta.select();
      document.execCommand("copy");
      ta.remove();
      setStatus("Copied");
      return true;
    } catch {
      setStatus("Copy failed");
      return false;
    }
  }

  // --- Detect context + UI update ---
  async function detect() {
    setStatus("Detecting…");
    try {
      const res = await sendCmd("DETECT");
      // Set manual item id if content script returns one
      if (res?.id != null) {
        $("#itemId").value = String(res.id);
      }
      // Build friendly context line
      const c = res?.ctx || {};
      const mode = res?.formMode || "unknown";
      const line = c.listTitle
        ? `Mode: ${mode} | Web: ${c.webUrl || "-"} | List: ${c.listTitle}`
        : `Mode: ${mode} | Web: ${c.webUrl || "-"}`;
      $("#ctxline").textContent = line;

      setStatus(res?.message || "Context detected");
    } catch (e) {
      console.error("[SP Reveal][popup] DETECT failed:", e);
      setStatus("Detect failed");
    }
  }

  // --- Run command for action buttons ---
  async function runAction(cmd) {
    setStatus("Working…");
    try {
      const res = await sendCmd(cmd);

      if (res?.copy != null) {
        await copyToClipboard(String(res.copy));
        // copyToClipboard already sets status to "Copied"
      } else {
        setStatus(res?.message || "Done");
      }

      // Optional behavior: close popup for SHOW_ALL_FIELDS so the top-layer dialog can take focus
      if (cmd === "SHOW_ALL_FIELDS") {
        setTimeout(() => window.close(), 120);
      }
    } catch (e) {
      console.error(`[SP Reveal][popup] ${cmd} failed:`, e);
      setStatus("Error – see console");
    }
  }

  // --- Manual ID setter ---
  async function applyManualId() {
    const raw = $("#itemId").value.trim();
    const id = raw ? Number(raw) : null;
    setStatus("Applying…");
    try {
      await sendCmd("SET_MANUAL_ID", {
        id: id != null && !Number.isNaN(id) ? id : null,
      });
      setStatus(
        id != null && !Number.isNaN(id) ? `Using ID ${id}` : "Auto ID mode",
      );
    } catch (e) {
      console.error("[SP Reveal][popup] SET_MANUAL_ID failed:", e);
      setStatus("Could not set ID");
    }
  }

  // --- Wire up events on load ---
  document.addEventListener("DOMContentLoaded", () => {
    // 1) Branding status on open (per Rik’s requirement)
    setStatus(BRANDING);

    // 2) Hook Detect button
    $("#btn-detect")?.addEventListener("click", detect);

    // 3) Hook manual ID apply
    $("#btn-insert-id")?.addEventListener("click", applyManualId);

    // 4) Hook all action buttons
    $all("button.action").forEach((btn) => {
      const cmd = btn?.dataset?.cmd;
      if (!cmd) return;
      btn.addEventListener("click", () => runAction(cmd));
    });

    // IMPORTANT: Do NOT auto-detect here, because that would immediately overwrite
    // the branding message with "Context detected". Detection should happen only
    // when the user clicks Detect or triggers an action requiring context.
  });
})();
