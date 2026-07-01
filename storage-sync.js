// ════════════════════════════════════════════════════════════════
// storage-sync.js — syncs uploaded Excel files through Supabase
// Slots map to your three existing file inputs:
//   inventory -> #fileInput        -> bucket path inventory/latest.xlsx
//   mapping   -> #mappingFileInput -> bucket path mapping/latest.xlsx
//   amc       -> #mosAmcFileInput  -> bucket path amc/latest.xlsx
//
// Runs AFTER auth.js, BEFORE/alongside script.js. Does not touch your
// existing parsing code — it reuses it by reading the file into the
// real <input type=file> and dispatching a "change" event, exactly
// as if the user picked the file themselves.
// ════════════════════════════════════════════════════════════════

const FILE_SLOTS = {
  inventory: { inputId: "fileInput",        path: "inventory/latest.xlsx" },
  mapping:   { inputId: "mappingFileInput", path: "mapping/latest.xlsx" },
  amc:       { inputId: "mosAmcFileInput",  path: "amc/latest.xlsx" },
};

const BUCKET = "inventory-files";

// ── Toast notifications ──
function ensureToastStyles() {
  if (document.getElementById("sync-toast-styles")) return;
  const style = document.createElement("style");
  style.id = "sync-toast-styles";
  style.textContent = `
    #sync-toast-stack {
      position: fixed; top: 16px; right: 16px; z-index: 10000;
      display: flex; flex-direction: column; gap: 8px; max-width: 320px;
    }
    .sync-toast {
      padding: 10px 14px; border-radius: 8px; font-size: 0.82rem;
      color: #fff; box-shadow: 0 4px 14px rgba(0,0,0,0.25);
      display: flex; align-items: center; gap: 8px;
      animation: sync-toast-in 0.2s ease-out;
    }
    .sync-toast.ok   { background: #1f8a4c; }
    .sync-toast.err  { background: #b3322a; }
    .sync-toast.info { background: #2563a8; }
    .sync-toast button {
      margin-left: auto; background: none; border: none; color: #fff;
      opacity: 0.8; cursor: pointer; font-size: 0.9rem; line-height: 1;
    }
    .sync-toast button:hover { opacity: 1; }
    @keyframes sync-toast-in { from { opacity: 0; transform: translateY(-6px); } to { opacity: 1; transform: translateY(0); } }
  `;
  document.head.appendChild(style);
}

function showToast(message, type = "info", timeoutMs = 4500) {
  ensureToastStyles();
  let stack = document.getElementById("sync-toast-stack");
  if (!stack) {
    stack = document.createElement("div");
    stack.id = "sync-toast-stack";
    document.body.appendChild(stack);
  }
  const el = document.createElement("div");
  el.className = `sync-toast ${type}`;
  el.innerHTML = `<span>${message}</span><button aria-label="Dismiss">✕</button>`;
  el.querySelector("button").addEventListener("click", () => el.remove());
  stack.appendChild(el);
  if (timeoutMs) setTimeout(() => el.remove(), timeoutMs);
  return el;
}

// ── Push: called when an admin selects a file ──
async function pushFileToSupabase(slot, file) {
  const { path } = FILE_SLOTS[slot];
  const sc = window.supabaseClient;
  const label = slot.charAt(0).toUpperCase() + slot.slice(1);

  const pending = showToast(`⏳ Syncing ${label} to Supabase…`, "info", 0);

  const { error: upErr } = await sc.storage.from(BUCKET).upload(path, file, {
    upsert: true,
    contentType: file.type || "application/octet-stream",
  });
  pending.remove();
  if (upErr) {
    console.error(`Upload failed (${slot}):`, upErr);
    showToast(`✗ ${label} sync failed: ${upErr.message || "unknown error"}`, "err");
    return;
  }

  const { error: metaErr } = await sc.from("app_files").upsert({
    slot,
    storage_path: path,
    filename: file.name,
    uploaded_by: window.APP_USER ? window.APP_USER.id : null,
    uploaded_at: new Date().toISOString(),
  });
  if (metaErr) {
    console.error(`Metadata save failed (${slot}):`, metaErr);
    showToast(`⚠️ ${label} file synced, but metadata save failed`, "err");
    return;
  }

  showToast(`✓ ${label} synced — all users will see this on refresh`, "ok");
}

// ── Pull: called on app load for every signed-in user ──
// Speed optimization: before downloading the (potentially large) Excel file,
// check the lightweight app_files metadata row for its uploaded_at timestamp.
// If we've already cached a blob for that exact timestamp (via the Cache API),
// reuse it instead of re-downloading — this is the difference between a few
// KB metadata check and a multi-MB file transfer on a slow connection.
const SYNC_CACHE_NAME = "inventory-files-v1";

async function pullFileFromSupabase(slot) {
  const { path, inputId } = FILE_SLOTS[slot];
  const sc = window.supabaseClient;

  let ts = null;
  try {
    const { data: meta } = await sc.from("app_files").select("uploaded_at").eq("slot", slot).maybeSingle();
    ts = meta && meta.uploaded_at ? meta.uploaded_at : null;
  } catch (e) {
    // Metadata check failed — fall through to a normal download below.
  }

  let blob = null;
  let cache = null;
  const cacheKey = ts ? `https://sync-cache.local/${slot}/${encodeURIComponent(ts)}` : null;

  if ("caches" in window && cacheKey) {
    try {
      cache = await caches.open(SYNC_CACHE_NAME);
      const cached = await cache.match(cacheKey);
      if (cached) blob = await cached.blob();
    } catch (e) {
      console.warn(`[storage-sync] Cache read failed (${slot}):`, e);
    }
  }

  if (!blob) {
    const { data, error } = await sc.storage.from(BUCKET).download(path);
    if (error) {
      // statusCode 404 / "Object not found" just means nothing uploaded yet — fine.
      // Anything else (permission denied, etc.) is worth knowing about.
      if (error.statusCode !== "404" && !/not.?found/i.test(error.message || "")) {
        console.error(`[storage-sync] Pull failed (${slot}):`, error);
      }
      return;
    }
    blob = data;
    if (cache && cacheKey) {
      try { await cache.put(cacheKey, new Response(blob)); } catch (e) { /* non-fatal */ }
    }
  }

  const filename = path.split("/").pop();
  const file = new File([blob], filename, { type: blob.type });

  const input = document.getElementById(inputId);
  if (!input) { console.warn(`[storage-sync] No input #${inputId} found for ${slot}`); return; }

  // Populate the real <input type=file> so existing handlers work unmodified.
  // Mark + unmark fromSupabase tightly around the synchronous dispatch so the
  // flag can never get stuck (previously it was set earlier in the caller and
  // only cleared by the listener — if the pull failed/returned early above,
  // it stayed stuck "on" forever and silently blocked all future real uploads).
  input.dataset.fromSupabase = "1";
  const dt = new DataTransfer();
  dt.items.add(file);
  input.files = dt.files;
  input.dispatchEvent(new Event("change", { bubbles: true })); // capture-phase listener consumes the flag synchronously here
  input.dataset.fromSupabase = "";
}

// ── Wire up: intercept admin's own file picks to also push to Supabase ──
// NOTE: script.js resets input.value = "" right after reading the file in
// its own (bubble-phase) change listener. We use { capture: true } here so
// OUR listener runs FIRST, before that reset wipes e.target.files.
function attachAdminUploadSync() {
  Object.entries(FILE_SLOTS).forEach(([slot, { inputId }]) => {
    const input = document.getElementById(inputId);
    if (!input) return;
    input.addEventListener(
      "change",
      (e) => {
        if (!window.isAdmin) return;            // safety net; UI is already hidden for non-admins
        const file = e.target.files && e.target.files[0];
        if (!file) { console.warn(`[storage-sync] No file found on change for ${slot}`); return; }
        // Don't re-upload the file we just pulled from Supabase ourselves
        if (input.dataset.fromSupabase === "1") { input.dataset.fromSupabase = ""; return; }
        pushFileToSupabase(slot, file);
      },
      { capture: true }
    );
  });
}

// ── Live update banner — shown to viewers (and admins) when someone
//    else uploads new data, via Supabase Realtime on app_files ──
function showNewDataBanner(slot) {
  if (document.getElementById("new-data-banner")) return; // already showing
  const label = slot.charAt(0).toUpperCase() + slot.slice(1);
  const el = document.createElement("div");
  el.id = "new-data-banner";
  el.style.cssText = `
    position: fixed; top: 0; left: 0; right: 0; z-index: 10001;
    background: var(--blue, #3a8fd4); color: #fff; text-align: center;
    padding: 10px 14px; font-size: 0.85rem;
    display: flex; align-items: center; justify-content: center; gap: 12px;
  `;
  el.innerHTML = `
    <span>🔔 New ${label} data is available.</span>
    <button id="refresh-data-btn" style="background:#fff;color:var(--blue,#3a8fd4);border:none;border-radius:6px;padding:4px 12px;font-weight:600;cursor:pointer;font-size:0.8rem">Refresh now</button>
    <button id="dismiss-banner-btn" style="background:none;border:1px solid rgba(255,255,255,0.6);color:#fff;border-radius:6px;padding:4px 10px;cursor:pointer;font-size:0.8rem">Later</button>
  `;
  document.body.prepend(el);
  document.getElementById("refresh-data-btn").addEventListener("click", () => location.reload());
  document.getElementById("dismiss-banner-btn").addEventListener("click", () => el.remove());
}

function attachRealtimeSync() {
  const sc = window.supabaseClient;
  sc.channel("app_files_changes")
    .on(
      "postgres_changes",
      { event: "*", schema: "public", table: "app_files" },
      (payload) => {
        const slot = payload.new && payload.new.slot;
        const uploadedBy = payload.new && payload.new.uploaded_by;
        // Don't show the banner to the admin who just uploaded it themselves
        if (window.APP_USER && uploadedBy === window.APP_USER.id) return;
        showNewDataBanner(slot || "inventory");
      }
    )
    .subscribe();
}

// ── On auth ready: load whatever's already in Supabase for everyone ──
document.addEventListener("epss-auth-ready", async () => {
  attachAdminUploadSync();
  attachRealtimeSync();

  // Run all three pulls concurrently instead of one-at-a-time — on a slow
  // connection this turns 3 sequential round-trips into 1.
  await Promise.all(Object.keys(FILE_SLOTS).map(slot => pullFileFromSupabase(slot)));
});
