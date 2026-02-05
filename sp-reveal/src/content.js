// ===== SP Reveal content script – modern SPO panel-aware + web scoping fix + ultra-resilient list detection =====

(() => {
  if (window.__spr_cs_loaded) return;
  window.__spr_cs_loaded = true;

  const STATE = {
    digest: null,
    digestTs: 0,
    manualId: null,
    listCache: null, // { listId, listTitle, listServerRelativeUrl, apiWebUrl }
  };

  // ---------- Core helpers ----------
  function pageCtxRaw() {
    return window._spPageContextInfo || {};
  }

  function getBasicCtx() {
    const raw = pageCtxRaw();
    return {
      origin: location.origin,
      webAbs: raw.webAbsoluteUrl || null,
      webSrvRel: raw.webServerRelativeUrl || null,
      listId: raw.listId ? String(raw.listId).replace(/[{}]/g, "") : null,
      listTitle: raw.listTitle || null,
    };
  }

  async function apiGet(apiWebUrl, rel) {
    const r = await fetch(`${apiWebUrl}${rel}`, {
      headers: { Accept: "application/json;odata=nometadata" },
      credentials: "same-origin",
    });
    if (!r.ok) throw new Error(`GET ${rel} → ${r.status}`);
    return r.json();
  }

  async function getDigest(apiWebUrl) {
    const now = Date.now();
    if (STATE.digest && now - STATE.digestTs < 9 * 60 * 1000)
      return STATE.digest;
    const r = await fetch(`${apiWebUrl}/_api/contextinfo`, {
      method: "POST",
      headers: { Accept: "application/json;odata=nometadata" },
      credentials: "same-origin",
    });
    const j = await r.json();
    STATE.digest = j.FormDigestValue;
    STATE.digestTs = Date.now();
    return STATE.digest;
  }

  async function apiPost(apiWebUrl, rel, body) {
    const digest = await getDigest(apiWebUrl);
    const r = await fetch(`${apiWebUrl}${rel}`, {
      method: "POST",
      headers: {
        Accept: "application/json;odata=nometadata",
        "Content-Type": "application/json;odata=nometadata",
        "X-RequestDigest": digest,
      },
      credentials: "same-origin",
      body: JSON.stringify(body || {}),
    });
    if (!r.ok) throw new Error(await r.text());
    return r.json();
  }

  // ---------- Paths & Web scoping ----------
  function deriveListServerRelativeUrl() {
    const p = decodeURIComponent(location.pathname);

    const lib = p.match(/^(.*?\/[^/]+)\/Forms\/AllItems\.aspx$/i);
    if (lib) return lib[1];

    const list = p.match(/^(.*?\/Lists\/[^/]+)\/[^/]+\.aspx$/i);
    if (list) return list[1];

    const classic = p.match(
      /^(.*?\/Lists\/[^/]+)\/(DispForm|EditForm|NewForm)\.aspx$/i,
    );
    if (classic) return classic[1];

    const qs = new URLSearchParams(location.search);
    const vp = qs.get("viewpath") || qs.get("RootFolder");
    if (vp) {
      const decoded = decodeURIComponent(vp);
      const m =
        decoded.match(/^(.*?\/Lists\/[^/]+)/i) ||
        decoded.match(/^(.*?\/[^/]+)(?:\/Forms)?\/AllItems\.aspx$/i);
      if (m) return m[1];
    }
    return null;
  }

  function deriveWebServerRelativeFromListPath(srvRel) {
    if (!srvRel) return null;
    const mList = srvRel.match(/^(.*)\/Lists\/[^/]+$/i);
    if (mList) return mList[1];
    const mLib = srvRel.match(/^(.*)\/[^/]+$/i);
    if (mLib) return mLib[1];
    return null;
  }

  function getApiWebUrl(basic, srvRelPath) {
    if (basic.webAbs) return basic.webAbs;
    if (basic.webSrvRel) return `${basic.origin}${basic.webSrvRel}`;
    const inferred = deriveWebServerRelativeFromListPath(srvRelPath);
    if (inferred) return `${basic.origin}${inferred}`;
    return basic.origin;
  }

  function extractListTitleFromPath(srvRel) {
    if (!srvRel) return null;
    const m = srvRel.match(/\/Lists\/([^/]+)$/i) || srvRel.match(/\/([^/]+)$/i);
    return m ? decodeURIComponent(m[1]) : null;
  }

  // ---------- DOM GUID probe ----------
  const GUID_RE =
    /[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}/g;

  function collectGuidCandidates() {
    const roots = [
      document,
      findModernFormDialog() || undefined,
      document.querySelector('[role="grid"]') || undefined,
      document.querySelector('[data-automationid="DetailsList"]') || undefined,
    ].filter(Boolean);

    const vals = new Set();
    const limitNodes = 800;

    for (const root of roots) {
      const walker = document.createTreeWalker(
        root,
        NodeFilter.SHOW_ELEMENT,
        null,
      );
      let count = 0;
      while (walker.nextNode() && count++ < limitNodes) {
        const el = walker.currentNode;
        for (const attr of el.attributes || []) {
          const m = String(attr.value).match(GUID_RE);
          if (m) m.forEach((g) => vals.add(g));
        }
      }
    }
    return Array.from(vals);
  }

  async function validateListGuid(apiWebUrl, guid) {
    try {
      const j = await apiGet(
        apiWebUrl,
        `/_api/web/lists(guid'${guid}')?$select=Id,Title,RootFolder/ServerRelativeUrl&$expand=RootFolder`,
      );
      const listId = (j.Id || "").replace(/[{}]/g, "");
      return {
        listId: listId || null,
        listTitle: j.Title || null,
        listServerRelativeUrl: j.RootFolder?.ServerRelativeUrl || null,
      };
    } catch {
      return null;
    }
  }

  // ---------- Modern modal/panel & item id ----------
  function findModernFormDialog() {
    const dialogs = Array.from(
      document.querySelectorAll(
        'div[role="dialog"], div[data-automationid="Panel"], div.ms-Layer, div[role="document"]',
      ),
    );
    const likely = dialogs.find((d) =>
      d.querySelector(
        '[data-automationid="FieldRenderer-label"], label.ms-Label',
      ),
    );
    return likely || dialogs[0] || null;
  }

  function extractIdFromUrl(u) {
    const keys = ["ID", "Id", "id", "ItemId", "item", "ItemID"];
    const qs = new URLSearchParams(u.search);
    for (const k of keys) {
      const v = qs.get(k);
      if (v && /^\d+$/.test(v)) return Number(v);
    }
    if (u.hash) {
      const hs = new URLSearchParams(u.hash.replace(/^#\??/, ""));
      for (const k of keys) {
        const v = hs.get(k);
        if (v && /^\d+$/.test(v)) return Number(v);
      }
    }
    return null;
  }

  function extractIdFromDialog(dialog) {
    if (!dialog) return null;
    const cands = [
      'input[name="ID"]',
      "input#ID",
      "[data-sp-item-id]",
      "[data-item-id]",
      "[data-listitemid]",
    ];
    for (const sel of cands) {
      const el = dialog.querySelector(sel);
      if (!el) continue;
      const v =
        el.value ||
        el.getAttribute("data-sp-item-id") ||
        el.getAttribute("data-item-id") ||
        el.getAttribute("data-listitemid");
      if (v && /^\d+$/.test(v)) return Number(v);
    }
    const href = dialog.querySelector('a[href*="ID="], a[href*="ItemId="]');
    if (href) {
      const u = new URL(href.href, location.origin);
      const idFromHref = extractIdFromUrl(u);
      if (idFromHref) return idFromHref;
    }
    return null;
  }

  function extractIdFromSelectedRow() {
    const row = document.querySelector(
      '[role="row"][aria-selected="true"], [data-selection-invoke="true"][aria-selected="true"]',
    );
    if (!row) return null;
    const cand =
      row.getAttribute("data-item-id") ||
      row.getAttribute("data-listitemid") ||
      row.getAttribute("data-selection-id") ||
      row.getAttribute("data-automationid");
    const m = cand && cand.match(/(\d+)/);
    return m ? Number(m[1]) : null;
  }

  function getFormMode() {
    const dlg = findModernFormDialog();
    if (dlg) {
      const aria = (dlg.getAttribute("aria-label") || "").toLowerCase();
      const text = (dlg.textContent || "").toLowerCase();
      if (aria.includes("new") || text.includes("new item")) return "new";
      if (aria.includes("edit") || text.includes("edit item")) return "edit";
      if (
        aria.includes("view") ||
        text.includes("view item") ||
        aria.includes("details")
      )
        return "view";
      return "dialog";
    }
    return "page";
  }

  async function detectItemContext() {
    const basic = getBasicCtx();
    const listPath = deriveListServerRelativeUrl();
    const apiWebUrl = getApiWebUrl(basic, listPath);

    // Manual override first
    let id = STATE.manualId || null;
    const u = new URL(location.href);
    if (!id) id = extractIdFromUrl(u);
    if (!id) id = extractIdFromDialog(findModernFormDialog());
    if (!id) id = extractIdFromSelectedRow();

    const listInfo = await resolveListInfoRobust({
      apiWebUrl,
      listPath,
      basic,
    });

    return {
      ok: true,
      ctx: {
        siteUrl: apiWebUrl,
        webUrl: apiWebUrl,
        listId: listInfo?.listId || basic.listId || null,
        listTitle: listInfo?.listTitle || basic.listTitle || null,
      },
      formMode: getFormMode(),
      id: id || null,
    };
  }

  async function resolveListInfoRobust({ apiWebUrl, listPath, basic }) {
    if (
      STATE.listCache &&
      STATE.listCache.apiWebUrl === apiWebUrl &&
      STATE.listCache.listServerRelativeUrl === listPath
    ) {
      return STATE.listCache;
    }

    const safePath = listPath
      ? listPath.startsWith("/")
        ? listPath
        : `/${listPath}`
      : null;
    const quoted = safePath ? `'${safePath.replace(/'/g, "''")}'` : null;

    if (basic.listId && basic.listTitle && safePath == null) {
      const v0 = {
        listId: basic.listId,
        listTitle: basic.listTitle,
        listServerRelativeUrl: null,
        apiWebUrl,
      };
      STATE.listCache = v0;
      return v0;
    }

    // 1) GetListUsingPath
    if (quoted) {
      try {
        const j = await apiGet(
          apiWebUrl,
          `/_api/web/GetListUsingPath(DecodedUrl=@a1)?@a1=${quoted}&$select=Id,Title,RootFolder/ServerRelativeUrl&$expand=RootFolder`,
        );
        const listId = (j.Id || "").replace(/[{}]/g, "");
        if (listId) {
          const v = {
            listId,
            listTitle: j.Title || null,
            listServerRelativeUrl: j.RootFolder?.ServerRelativeUrl || safePath,
            apiWebUrl,
          };
          STATE.listCache = v;
          return v;
        }
      } catch {}
    }

    // 2) GetList
    if (quoted) {
      try {
        const j = await apiGet(
          apiWebUrl,
          `/_api/web/GetList(@a1)?@a1=${quoted}&$select=Id,Title,RootFolder/ServerRelativeUrl&$expand=RootFolder`,
        );
        const listId = (j.Id || "").replace(/[{}]/g, "");
        if (listId) {
          const v = {
            listId,
            listTitle: j.Title || null,
            listServerRelativeUrl: j.RootFolder?.ServerRelativeUrl || safePath,
            apiWebUrl,
          };
          STATE.listCache = v;
          return v;
        }
      } catch {}
    }

    // 3) Folder → ParentList
    if (quoted) {
      try {
        const f = await apiGet(
          apiWebUrl,
          `/_api/web/GetFolderByServerRelativeUrl(${quoted})?$select=ServerRelativeUrl,ListItemAllFields/ParentList/Id,ListItemAllFields/ParentList/Title&$expand=ListItemAllFields/ParentList`,
        );
        const id = (f?.ListItemAllFields?.ParentList?.Id || "").replace(
          /[{}]/g,
          "",
        );
        if (id) {
          const v = {
            listId: id,
            listTitle: f?.ListItemAllFields?.ParentList?.Title || null,
            listServerRelativeUrl: f?.ServerRelativeUrl || safePath,
            apiWebUrl,
          };
          STATE.listCache = v;
          return v;
        }
      } catch {}
    }

    // 4) File(…/AllItems.aspx) → ParentList
    if (safePath) {
      const fileUrl = safePath.endsWith("/AllItems.aspx")
        ? safePath
        : `${safePath}/AllItems.aspx`;
      const q = `'${fileUrl.replace(/'/g, "''")}'`;
      try {
        const f = await apiGet(
          apiWebUrl,
          `/_api/web/GetFileByServerRelativeUrl(${q})/ListItemAllFields/ParentList?$select=Id,Title,RootFolder/ServerRelativeUrl&$expand=RootFolder`,
        );
        const id = (f?.Id || "").replace(/[{}]/g, "");
        if (id) {
          const v = {
            listId: id,
            listTitle: f?.Title || null,
            listServerRelativeUrl: f?.RootFolder?.ServerRelativeUrl || safePath,
            apiWebUrl,
          };
          STATE.listCache = v;
          return v;
        }
      } catch {}
    }

    // 5) GetByTitle
    const maybeTitle = extractListTitleFromPath(safePath);
    if (maybeTitle) {
      const qTitle = `'${maybeTitle.replace(/'/g, "''")}'`;
      try {
        const j = await apiGet(
          apiWebUrl,
          `/_api/web/Lists/GetByTitle(${qTitle})?$select=Id,Title,RootFolder/ServerRelativeUrl&$expand=RootFolder`,
        );
        const listId = (j.Id || "").replace(/[{}]/g, "");
        if (listId) {
          const v = {
            listId,
            listTitle: j.Title || maybeTitle,
            listServerRelativeUrl: j.RootFolder?.ServerRelativeUrl || safePath,
            apiWebUrl,
          };
          STATE.listCache = v;
          return v;
        }
      } catch {}
    }

    // 6) Enumerate all lists
    try {
      const all = await apiGet(
        apiWebUrl,
        `/_api/web/lists?$select=Id,Title,RootFolder/ServerRelativeUrl&$expand=RootFolder&$top=5000`,
      );
      const items = Array.isArray(all?.value) ? all.value : [];
      let found = null;
      if (safePath) {
        found = items.find(
          (x) =>
            (x.RootFolder?.ServerRelativeUrl || "").toLowerCase() ===
            safePath.toLowerCase(),
        );
      }
      if (!found && maybeTitle) {
        found = items.find(
          (x) => (x.Title || "").toLowerCase() === maybeTitle.toLowerCase(),
        );
      }
      if (found) {
        const listId = (found.Id || "").replace(/[{}]/g, "");
        const v = {
          listId,
          listTitle: found.Title || null,
          listServerRelativeUrl: found.RootFolder?.ServerRelativeUrl || null,
          apiWebUrl,
        };
        STATE.listCache = v;
        return v;
      }
    } catch {}

    // 7) DOM GUID probe
    const candidates = collectGuidCandidates();
    for (const guid of candidates) {
      const v = await validateListGuid(apiWebUrl, guid);
      if (v?.listId) {
        STATE.listCache = { ...v, apiWebUrl };
        return STATE.listCache;
      }
    }

    return null;
  }

  // ---------- UX helpers: clipboard + toast ----------
  async function copyToClipboard(text) {
    const str = String(text ?? "");
    try {
      if (navigator.clipboard && window.isSecureContext) {
        await navigator.clipboard.writeText(str);
        return true;
      }
      // Fallback
      const ta = document.createElement("textarea");
      ta.value = str;
      ta.style.position = "fixed";
      ta.style.top = "-9999px";
      document.body.appendChild(ta);
      ta.focus();
      ta.select();
      const ok = document.execCommand("copy");
      ta.remove();
      return ok;
    } catch {
      return false;
    }
  }

  function getToast() {
    let t = document.getElementById("spr-toast");
    if (!t) {
      t = document.createElement("div");
      t.id = "spr-toast";
      t.setAttribute("role", "status");
      t.setAttribute("aria-live", "polite");
      // Minimal inline style in case content.css is not injected
      t.style.cssText =
        "position:fixed;left:50%;bottom:24px;transform:translateX(-50%);background:#323130;color:#fff;padding:8px 12px;border-radius:6px;z-index:2147483647;display:none;font:12px Segoe UI, Arial, sans-serif;";
      document.documentElement.appendChild(t);
    }
    return t;
  }

  let toastTimer = null;
  function showToast(msg = "Copied") {
    const t = getToast();
    t.textContent = msg;
    t.style.display = "block";
    if (toastTimer) clearTimeout(toastTimer);
    toastTimer = setTimeout(() => {
      t.style.display = "none";
    }, 1200);
  }

  function wireCopyBehavior(el, value) {
    if (!el || el.dataset.sprWired === "1") return;
    el.dataset.sprWired = "1";
    el.setAttribute("role", "button");
    el.setAttribute("tabindex", "0");
    el.setAttribute("title", "Click to copy");
    // Ensure pointer cursor even if stylesheet isn't injected
    el.style.cursor = "pointer";

    const handler = async (e) => {
      e.preventDefault();
      e.stopPropagation();
      const ok = await copyToClipboard(value);
      showToast(ok ? `Copied: ${value}` : "Copy failed");
      // subtle visual feedback
      el.style.outline = "2px solid #0078d4";
      setTimeout(() => (el.style.outline = ""), 250);
    };

    el.addEventListener("click", handler);
    el.addEventListener("keydown", (e) => {
      if (e.key === "Enter" || e.key === " ") handler(e);
    });
  }

  // ---------- Features ----------
  function escapeHtml(s) {
    return String(s).replace(
      /[&<>"']/g,
      (ch) =>
        ({
          "&": "&amp;",
          "<": "&lt;",
          ">": "&gt;",
          '"': "&quot;",
          "'": "&#39;",
        })[ch],
    );
  }

  function highlightRaw(raw, query) {
    if (!query) return escapeHtml(String(raw));
    const q = String(query);
    const esc = q.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const re = new RegExp(esc, "ig");
    const str = String(raw);
    let out = "";
    let last = 0;
    let m;
    while ((m = re.exec(str)) !== null) {
      out +=
        escapeHtml(str.slice(last, m.index)) +
        "<mark>" +
        escapeHtml(m[0]) +
        "</mark>";
      last = m.index + m[0].length;
    }
    out += escapeHtml(str.slice(last));
    return out;
  }

  async function ensureListAndId() {
    const d = await detectItemContext();
    if (!d.ctx.listId) return { error: "No list context detected" };
    if (!d.id)
      return {
        error: "Could not determine Item ID (for New forms, save first)",
      };
    return d;
  }

  async function showLogicalNames() {
    const d = await detectItemContext();
    if (!d.ctx.listId) return { message: "No list context detected" };

    const fields = await apiGet(
      d.ctx.webUrl,
      `/_api/web/lists(guid'${d.ctx.listId}')/fields?$select=Title,InternalName`,
    );
    const map = new Map(fields.value.map((f) => [f.Title, f.InternalName]));

    const root = findModernFormDialog() || document;
    const sels = [
      '[data-automationid="FieldRenderer-label"]',
      "label.ms-Label",
      'div[role="heading"]',
      'label[for^="Field"]',
    ];
    sels.forEach((sel) => {
      root.querySelectorAll(sel).forEach((el) => {
        const title = (el.textContent || "").trim();
        const internal = title && map.get(title);
        if (internal) {
          let tag = el.querySelector(".spr-internal-tag");
          if (!tag) {
            tag = document.createElement("span");
            tag.className = "spr-internal-tag";
            tag.textContent = internal;
            // inline style to ensure visibility even without stylesheet
            tag.style.cssText =
              "color:#6b6b6b;font-weight:600;margin-left:6px;font-size:11px;background:#f3f2f1;border:1px solid #e1dfdd;padding:0 4px;border-radius:4px;cursor:pointer;";
            el.appendChild(tag);
          } else {
            // keep label in sync if field map changed
            tag.textContent = internal;
          }
          // Make it clickable & copy-ready
          wireCopyBehavior(tag, internal);
        }
      });
    });
    return { message: "Logical names shown (click a tag to copy)" };
  }

  function clearLogicalNames() {
    document.querySelectorAll(".spr-internal-tag").forEach((n) => n.remove());
    return { message: "Cleared" };
  }

  async function buildItemUrl() {
    const d = await ensureListAndId();
    if (d.error) return { message: d.error };

    // Get the default display form (server-relative in SPO)
    const list = await apiGet(
      d.ctx.webUrl,
      `/_api/web/lists(guid'${d.ctx.listId}')?$select=DefaultDisplayFormUrl`,
    );

    const display = list?.DefaultDisplayFormUrl || "";
    const isAbs = /^https?:\/\//i.test(display);
    const origin = new URL(d.ctx.webUrl).origin;

    // Normalize to absolute URL using origin to avoid duplicating the site path
    const baseHref = isAbs ? display : `${origin}${display}`;

    // Use URL API to safely add/set the ID query parameter
    const u = new URL(baseHref, origin);
    u.searchParams.set("ID", String(d.id));

    return {
      copy: u.toString(),
      message: "Item URL ready",
    };
  }

  async function buildApiUrl() {
    const d = await ensureListAndId();
    if (d.error) return { message: d.error };
    return {
      copy: `${d.ctx.webUrl}/_api/web/lists(guid'${d.ctx.listId}')/items(${d.id})`,
      message: "API URL ready",
    };
  }

  async function getItemIdToCopy() {
    const d = await ensureListAndId();
    if (d.error) return { message: d.error };
    return { copy: String(d.id), message: "Item ID ready" };
  }

  async function duplicateItem() {
    const d = await ensureListAndId();
    if (d.error) return { message: d.error };

    const meta = await apiGet(
      d.ctx.webUrl,
      `/_api/web/lists(guid'${d.ctx.listId}')/fields?$select=InternalName,TypeAsString,ReadOnlyField,Hidden,Sealed`,
    );
    const allow = new Set([
      "Text",
      "Note",
      "Number",
      "Currency",
      "DateTime",
      "Boolean",
      "Choice",
    ]);
    const writable = new Set(
      meta.value
        .filter(
          (f) =>
            !f.ReadOnlyField &&
            !f.Hidden &&
            !f.Sealed &&
            allow.has(f.TypeAsString),
        )
        .map((f) => f.InternalName),
    );

    const src = await apiGet(
      d.ctx.webUrl,
      `/_api/web/lists(guid'${d.ctx.listId}')/items(${d.id})?$select=*`,
    );
    const payload = {};
    Object.keys(src).forEach((k) => {
      if (writable.has(k)) payload[k] = src[k];
    });
    if (src.ContentTypeId) payload["ContentTypeId"] = src.ContentTypeId;

    const created = await apiPost(
      d.ctx.webUrl,
      `/_api/web/lists(guid'${d.ctx.listId}')/items`,
      payload,
    );
    const list = await apiGet(
      d.ctx.webUrl,
      `/_api/web/lists(guid'${d.ctx.listId}')?$select=DefaultDisplayFormUrl`,
    );
    const origin = new URL(d.ctx.webUrl).origin;
    const baseHref = /^https?:\/\//i.test(list.DefaultDisplayFormUrl)
      ? list.DefaultDisplayFormUrl
      : `${origin}${list.DefaultDisplayFormUrl}`;
    const newUrl = new URL(baseHref, origin);
    newUrl.searchParams.set("ID", String(created.Id));

    return {
      copy: newUrl.toString(),
      message: `Duplicated as ID ${created.Id}`,
    };
  }

  // --- New: list-level helpers ---
  async function copyListGuid() {
    const d = await detectItemContext();
    const listId = d?.ctx?.listId;
    if (!listId) return { message: "No list context detected" };
    return { copy: String(listId), message: "List GUID ready" };
  }

  async function copyListUrl() {
    const d = await detectItemContext();
    const listId = d?.ctx?.listId;
    if (!listId) return { message: "No list context detected" };
    const j = await apiGet(
      d.ctx.webUrl,
      `/_api/web/lists(guid'${listId}')?$select=RootFolder/ServerRelativeUrl&$expand=RootFolder`,
    );
    const serverRel = j?.RootFolder?.ServerRelativeUrl;
    if (!serverRel) return { message: "Could not resolve list URL" };
    const origin = new URL(d.ctx.webUrl).origin;
    return { copy: `${origin}${serverRel}`, message: "List URL ready" };
  }

  async function copyListApiUrl() {
    const d = await detectItemContext();
    const listId = d?.ctx?.listId;
    if (!listId) return { message: "No list context detected" };
    return {
      copy: `${d.ctx.webUrl}/_api/web/lists(guid'${listId}')`,
      message: "List API URL ready",
    };
  }

  // ---------- SPO form/panel helpers (detect, close robustly, reopen without new tabs) ----------
  function getOpenPanel() {
    const candidates = Array.from(
      document.querySelectorAll(
        'div[data-automationid="Panel"], div[role="dialog"]',
      ),
    );
    for (const p of candidates) {
      const cs = getComputedStyle(p);
      const hidden =
        cs.display === "none" ||
        cs.visibility === "hidden" ||
        parseFloat(cs.opacity || "1") < 0.05;
      const rect = p.getBoundingClientRect();
      if (hidden || rect.width < 40 || rect.height < 40) continue;
      if (rect.bottom < 0 || rect.top > (window.innerHeight || 800)) continue;

      const closeBtn =
        p.querySelector('[data-automationid="PanelCloseButton"]') ||
        p.querySelector('button[aria-label="Close"]') ||
        p.querySelector('button[title="Close"]') ||
        p.querySelector('button[aria-label*="Close"]') ||
        p.querySelector('button[aria-label="Sluiten"]') ||
        p.querySelector('button[title="Sluiten"]') ||
        p.querySelector('button[aria-label*="Sluit"]') ||
        p.querySelector('button[aria-label*="Dismiss"]');

      return { panel: p, closeBtn };
    }
    return null;
  }

  function forceSelfAndClick(el) {
    if (!el) return;
    try {
      el.setAttribute("target", "_self");
    } catch {}
    el.dispatchEvent(
      new MouseEvent("pointerdown", { bubbles: true, cancelable: true }),
    );
    el.dispatchEvent(
      new MouseEvent("mousedown", { bubbles: true, cancelable: true }),
    );
    el.dispatchEvent(
      new MouseEvent("mouseup", { bubbles: true, cancelable: true }),
    );
    el.dispatchEvent(
      new MouseEvent("click", { bubbles: true, cancelable: true }),
    );
  }

  function dispatchEscape(target) {
    const opts = {
      key: "Escape",
      code: "Escape",
      keyCode: 27,
      which: 27,
      bubbles: true,
      cancelable: true,
    };
    target.dispatchEvent(new KeyboardEvent("keydown", opts));
    target.dispatchEvent(new KeyboardEvent("keyup", opts));
  }

  async function waitFor(cond, timeoutMs = 2000, stepMs = 60) {
    const start = Date.now();
    return new Promise((resolve) => {
      (function tick() {
        if (cond()) return resolve(true);
        if (Date.now() - start >= timeoutMs) return resolve(false);
        setTimeout(tick, stepMs);
      })();
    });
  }

  async function tryHistoryBackClose() {
    const before = location.href;
    try {
      history.back();
      const ok = await waitFor(
        () =>
          !document.querySelector(
            'button.sp-itemDialog-closeBtn, .sp-itemDialog-closeBtn, [class*="sp-itemDialog-closeBtn"]',
          ) && !getOpenPanel(),
        800,
        50,
      );
      if (!ok && location.href !== before) {
        try {
          history.forward();
        } catch {}
      }
      return ok;
    } catch {
      return false;
    }
  }

  async function closeSharePointFormIfOpen() {
    // Prefer the modern item dialog close button if present (tenant-specific class)
    const itemDlgBtn = document.querySelector(
      'button.sp-itemDialog-closeBtn, .sp-itemDialog-closeBtn, [class*="sp-itemDialog-closeBtn"], button.ms-CommandBarItem-link.sp-itemDialog-closeBtn',
    );
    if (itemDlgBtn) {
      forceSelfAndClick(itemDlgBtn);
      if (
        await waitFor(
          () =>
            !document.querySelector(
              'button.sp-itemDialog-closeBtn, .sp-itemDialog-closeBtn, [class*="sp-itemDialog-closeBtn"]',
            ),
          800,
          40,
        )
      )
        return true;

      // ESC fallback
      try {
        const active = document.activeElement || document.body;
        for (let i = 0; i < 2; i++) {
          dispatchEscape(active);
          dispatchEscape(document);
        }
      } catch {}
      if (
        await waitFor(
          () =>
            !document.querySelector(
              'button.sp-itemDialog-closeBtn, .sp-itemDialog-closeBtn, [class*="sp-itemDialog-closeBtn"]',
            ),
          600,
          40,
        )
      )
        return true;

      // Overlay fallback
      try {
        const overlays = Array.from(
          document.querySelectorAll(
            '.ms-Overlay, div[role="presentation"], .ms-Layer',
          ),
        );
        const overlay = overlays.find((o) => {
          const r = o.getBoundingClientRect();
          const cs = getComputedStyle(o);
          return (
            r.width > 40 &&
            r.height > 40 &&
            cs.opacity !== "0" &&
            cs.visibility !== "hidden"
          );
        });
        if (overlay) forceSelfAndClick(overlay);
      } catch {}
      if (
        await waitFor(
          () =>
            !document.querySelector(
              'button.sp-itemDialog-closeBtn, .sp-itemDialog-closeBtn, [class*="sp-itemDialog-closeBtn"]',
            ),
          600,
          40,
        )
      )
        return true;

      if (await tryHistoryBackClose()) return true;
    }

    // Fluent Panel fallback paths
    const found = getOpenPanel();
    if (found) {
      if (found.closeBtn) {
        forceSelfAndClick(found.closeBtn);
        if (await waitFor(() => !getOpenPanel(), 800, 40)) return true;
      }
      try {
        const active = document.activeElement || document.body;
        for (let i = 0; i < 2; i++) {
          dispatchEscape(active);
          dispatchEscape(found.panel);
          dispatchEscape(document);
        }
      } catch {}
      if (await waitFor(() => !getOpenPanel(), 600, 40)) return true;

      try {
        const overlays = Array.from(
          document.querySelectorAll(
            '.ms-Overlay, div[role="presentation"], .ms-Layer',
          ),
        );
        const overlay = overlays.find((o) => {
          const r = o.getBoundingClientRect();
          const cs = getComputedStyle(o);
          return (
            r.width > 40 &&
            r.height > 40 &&
            cs.opacity !== "0" &&
            cs.visibility !== "hidden"
          );
        });
        if (overlay) forceSelfAndClick(overlay);
      } catch {}
      if (await waitFor(() => !getOpenPanel(), 600, 40)) return true;

      if (await tryHistoryBackClose()) return true;
    }

    return true; // best effort
  }

  function buildReopenHandle(itemId) {
    const idStr = String(itemId);
    const grid =
      document.querySelector('[role="grid"]') ||
      document.querySelector('[data-automationid="DetailsList"]') ||
      document;
    const selectedRow = grid.querySelector(
      '[role="row"][aria-selected="true"], [data-selection-invoke="true"][aria-selected="true"]',
    );

    const linkSelectors = [
      `a[href*="ID=${idStr}"]`,
      `a[href*="ItemId=${idStr}"]`,
      'a[data-automationid="FieldRenderer-title"]',
      "a.ms-DetailsRow-cellLink",
    ];

    if (selectedRow) {
      for (const sel of linkSelectors) {
        const a = selectedRow.querySelector(sel);
        if (a && a.href) {
          return () => {
            try {
              forceSelfAndClick(a);
            } catch {}
          };
        }
      }
      return () => {
        try {
          forceSelfAndClick(selectedRow);
        } catch {}
      };
    }

    for (const sel of linkSelectors) {
      const a = grid.querySelector(sel);
      if (a && a.href) {
        return () => {
          try {
            forceSelfAndClick(a);
          } catch {}
        };
      }
    }
    return null; // avoid new tab fallbacks
  }

  // ---------- Show all fields: close SPO form, themed dialog (Fluent light/dark), unified typography, search, reopen on close ----------
  async function showAllFields() {
    const d = await ensureListAndId();
    if (d.error) return { message: d.error };

    const reopen = buildReopenHandle(d.id);

    await closeSharePointFormIfOpen();

    const item = await apiGet(
      d.ctx.webUrl,
      `/_api/web/lists(guid'${d.ctx.listId}')/items(${d.id})?$select=*`,
    );
    const all = Object.keys(item).map((k) => {
      const v = item[k];
      const pretty =
        typeof v === "object" ? JSON.stringify(v, null, 2) : String(v);
      return {
        key: k,
        pretty,
        keyL: k.toLowerCase(),
        prettyL: pretty.toLowerCase(),
      };
    });
    const json = JSON.stringify(item, null, 2);

    // Inject themed styles once (matches popup look & feel + typography)
    let style = document.getElementById("spr-dialog-style");
    if (!style) {
      style = document.createElement("style");
      style.id = "spr-dialog-style";
      style.textContent = `
        /* Theme tokens scoped to dialog */
        #spr-dialog {
          --sp-blue: #0078d4;
          --sp-teal: #00a5a3;

          /* Light theme defaults */
          --theme-bg: #ffffff;
          --theme-header-bg: #f3f2f1;
          --theme-border: #e1dfdd;
          --theme-text: #242424;
          --theme-muted: #666666;
          --theme-button-bg: #ffffff;
          --theme-button-border: #e1dfdd;
          --theme-button-hover: #f7fbff;
          --theme-divider: #f1f1f1;
        }
        @media (prefers-color-scheme: dark) {
          #spr-dialog {
            --theme-bg: #1f1f1f;
            --theme-header-bg: #2a2a2a;
            --theme-border: #3c3c3c;
            --theme-text: #ffffff;
            --theme-muted: #bbbbbb;
            --theme-button-bg: #292929;
            --theme-button-border: #3c3c3c;
            --theme-button-hover: #3a3f42;
            --theme-divider: #3a3a3a;
          }
        }

        /* ==== UNIFIED FLUENT TYPOGRAPHY (matching popup) ==== */
        #spr-dialog,
        #spr-dialog * {
          font-family: "Segoe UI", Arial, sans-serif;
          font-size: 13px;
          line-height: 1.4;
        }
        #spr-dialog .spr-head {
          font-size: 14px;
          font-weight: 600;
        }
        #spr-searchbar input[type="search"] {
          font-size: 13px;
        }
        #spr-searchbar .spr-count {
          font-size: 12px;
        }
        #spr-dialog table {
          font-size: 12px;
        }
        #spr-dialog th {
          font-size: 12px;
          font-weight: 600;
        }
        #spr-dialog .spr-btn {
          font-size: 13px;
        }

        #spr-dialog { border: 0; padding: 0; border-radius: 10px; color: var(--theme-text); }
        #spr-dialog::backdrop { background: rgba(0,0,0,.35); }

        /* Card container (opaque) */
        #spr-dialog .spr-card {
          display:flex; flex-direction:column; max-height:80vh;
          background: var(--theme-bg);
          border: 1px solid var(--theme-border);
          box-shadow: 0 12px 28px rgba(0,0,0,.2);
          border-radius: 10px;
          color: var(--theme-text);
        }

        /* Header */
        #spr-dialog .spr-head {
          padding: 10px 12px;
          background: var(--theme-header-bg);
          border-bottom: 1px solid var(--theme-border);
        }

        /* Search bar */
        #spr-searchbar {
          padding: 8px 12px;
          border-bottom: 1px solid var(--theme-border);
          display: flex; align-items: center; gap: 8px;
          background: var(--theme-header-bg);
          color: var(--theme-text);
        }
        #spr-searchbar input[type="search"] {
          flex: 1;
          padding: 8px 10px;
          border: 1px solid var(--theme-border);
          border-radius: 6px;
          background: var(--theme-bg);
          color: var(--theme-text);
        }
        #spr-searchbar .spr-count {
          color: var(--theme-muted);
        }

        /* Content */
        #spr-dialog .spr-content { padding: 12px; overflow: auto; }
        #spr-dialog table { border-collapse: collapse; width: 100%; }
        #spr-dialog th, #spr-dialog td { border-bottom: 1px solid var(--theme-divider); padding: 6px 8px; vertical-align: top; }
        #spr-dialog th { text-align: left; width: 260px; background: var(--theme-header-bg); }
        #spr-dialog pre { margin: 0; white-space: pre-wrap; color: var(--theme-text); }
        #spr-dialog mark { background: #fff59d; padding: 0 .5px; }

        /* Footer & buttons */
        #spr-dialog .spr-foot {
          padding: 8px 12px;
          border-top: 1px solid var(--theme-border);
          display: flex; justify-content: flex-end; gap: 8px;
          background: var(--theme-bg);
        }
        #spr-dialog .spr-btn {
          padding: 6px 10px;
          border: 1px solid var(--theme-button-border);
          border-radius: 6px;
          background: var(--theme-button-bg);
          color: var(--theme-text);
          cursor: pointer;
        }
        #spr-dialog .spr-btn:hover {
          background: var(--theme-button-hover);
          border-color: var(--sp-blue);
        }
      `;
      document.head.appendChild(style);
    }

    // Ensure single dialog
    let dlg = document.getElementById("spr-dialog");
    if (dlg) {
      try {
        dlg.close();
      } catch {}
      dlg.remove();
    }

    // Build dialog
    dlg = document.createElement("dialog");
    dlg.id = "spr-dialog";
    dlg.setAttribute("aria-label", "All fields");
    dlg.style.width = "min(900px, 92vw)";
    dlg.style.maxHeight = "80vh";
    dlg.style.overflow = "hidden";

    const card = document.createElement("div");
    card.className = "spr-card";

    const head = document.createElement("div");
    head.className = "spr-head";
    head.textContent = "All fields";

    // Search bar
    const searchbar = document.createElement("div");
    searchbar.id = "spr-searchbar";

    const searchInput = document.createElement("input");
    searchInput.type = "search";
    searchInput.placeholder = "Search fields or values…";
    searchInput.autocomplete = "off";
    searchInput.spellcheck = false;

    const countSpan = document.createElement("span");
    countSpan.className = "spr-count";
    countSpan.textContent = `${all.length} of ${all.length}`;

    searchbar.append(searchInput, countSpan);

    const contentEl = document.createElement("div");
    contentEl.className = "spr-content";

    const footer = document.createElement("div");
    footer.className = "spr-foot";

    const closeBtn = document.createElement("button");
    closeBtn.className = "spr-btn";
    closeBtn.id = "spr-close";
    closeBtn.type = "button";
    closeBtn.textContent = "Close";

    const copyBtn = document.createElement("button");
    copyBtn.className = "spr-btn";
    copyBtn.id = "spr-copy";
    copyBtn.type = "button";
    copyBtn.textContent = "Copy JSON";

    footer.append(closeBtn, copyBtn);
    card.append(head, searchbar, contentEl, footer);
    dlg.appendChild(card);
    document.documentElement.appendChild(dlg);

    // Rendering with filtering + highlighting
    function render(query) {
      const q = (query || "").trim().toLowerCase();
      const rows = q
        ? all.filter((r) => r.keyL.includes(q) || r.prettyL.includes(q))
        : all;
      countSpan.textContent = `${rows.length} of ${all.length}`;

      const htmlRows = rows
        .map((r) => {
          const keyHtml = highlightRaw(r.key, q);
          const valHtml = highlightRaw(r.pretty, q);
          return `<tr>
          <th>${keyHtml}</th>
          <td><pre>${valHtml}</pre></td>
        </tr>`;
        })
        .join("");

      contentEl.innerHTML = `<table><tbody>${htmlRows}</tbody></table>`;
    }

    render("");

    // Debounced input listener
    let t = null;
    searchInput.addEventListener("input", () => {
      if (t) clearTimeout(t);
      const v = searchInput.value;
      t = setTimeout(() => render(v), 120);
    });

    // Keep full JSON available for copy
    contentEl.setAttribute("data-json", json);

    function cleanupAndReopen() {
      try {
        dlg.close();
      } catch {}
      dlg.remove();
      if (typeof reopen === "function") {
        setTimeout(() => {
          try {
            reopen();
          } catch {}
        }, 60);
      }
    }

    closeBtn.addEventListener("click", (e) => {
      e.preventDefault();
      cleanupAndReopen();
    });

    copyBtn.addEventListener("click", async (e) => {
      e.preventDefault();
      try {
        const payload = contentEl.getAttribute("data-json") || "{}";
        await navigator.clipboard.writeText(payload);
        const old = copyBtn.textContent;
        copyBtn.textContent = "Copied!";
        setTimeout(() => (copyBtn.textContent = old), 900);
      } catch {
        const old = copyBtn.textContent;
        copyBtn.textContent = "Copy failed";
        setTimeout(() => (copyBtn.textContent = old), 1200);
      }
    });

    dlg.addEventListener("cancel", (e) => {
      e.preventDefault();
      cleanupAndReopen();
    });

    dlg.showModal();
    dlg.focus();
    // Auto-focus search for instant typing
    setTimeout(() => searchInput.focus(), 0);

    return { message: "Fields shown" };
  }

  // ---------- Messaging ----------
  chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
    (async () => {
      try {
        switch (msg.cmd) {
          case "DETECT": {
            const d = await detectItemContext();
            sendResponse({ ok: true, ...d });
            break;
          }
          case "SET_MANUAL_ID":
            STATE.manualId = msg.args?.id || null;
            sendResponse({ ok: true });
            break;

          case "SHOW_INTERNAL":
            sendResponse(await showLogicalNames());
            break;
          case "CLEAR_INTERNAL":
            sendResponse(clearLogicalNames());
            break;

          case "COPY_ITEM_URL":
            sendResponse(await buildItemUrl());
            break;
          case "COPY_API_URL":
            sendResponse(await buildApiUrl());
            break;
          case "COPY_ITEM_ID":
            sendResponse(await getItemIdToCopy());
            break;
          case "DUPLICATE":
            sendResponse(await duplicateItem());
            break;
          case "SHOW_ALL_FIELDS":
            sendResponse(await showAllFields());
            break;

          // List-level routes
          case "COPY_LIST_GUID":
            sendResponse(await copyListGuid());
            break;
          case "COPY_LIST_URL":
            sendResponse(await copyListUrl());
            break;
          case "COPY_LIST_API_URL":
            sendResponse(await copyListApiUrl());
            break;

          default:
            sendResponse({ ok: false, message: "Unknown command" });
        }
      } catch (e) {
        console.error(e);
        sendResponse({ ok: false, message: e?.message || "Error" });
      }
    })();
    return true;
  });
})();
