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

// ── Push: called when an admin selects a file ──
async function pushFileToSupabase(slot, file) {
  const { path } = FILE_SLOTS[slot];
  const sc = window.supabaseClient;

  const { error: upErr } = await sc.storage.from(BUCKET).upload(path, file, {
    upsert: true,
    contentType: file.type || "application/octet-stream",
  });
  if (upErr) { console.error(`Upload failed (${slot}):`, upErr); return; }

  const { error: metaErr } = await sc.from("app_files").upsert({
    slot,
    storage_path: path,
    filename: file.name,
    uploaded_by: window.APP_USER ? window.APP_USER.id : null,
    uploaded_at: new Date().toISOString(),
  });
  if (metaErr) console.error(`Metadata save failed (${slot}):`, metaErr);
}

// ── Pull: called on app load for every signed-in user ──
async function pullFileFromSupabase(slot) {
  const { path, inputId } = FILE_SLOTS[slot];
  const sc = window.supabaseClient;

  const { data: blob, error } = await sc.storage.from(BUCKET).download(path);
  if (error) {
    // Nothing uploaded yet for this slot — that's fine, not an error state
    return;
  }

  const filename = path.split("/").pop();
  const file = new File([blob], filename, { type: blob.type });

  const input = document.getElementById(inputId);
  if (!input) return;

  // Populate the real <input type=file> so existing handlers work unmodified
  const dt = new DataTransfer();
  dt.items.add(file);
  input.files = dt.files;
  input.dispatchEvent(new Event("change", { bubbles: true }));
}

// ── Wire up: intercept admin's own file picks to also push to Supabase ──
function attachAdminUploadSync() {
  Object.entries(FILE_SLOTS).forEach(([slot, { inputId }]) => {
    const input = document.getElementById(inputId);
    if (!input) return;
    input.addEventListener("change", (e) => {
      if (!window.isAdmin) return;            // safety net; UI is already hidden for non-admins
      const file = e.target.files && e.target.files[0];
      if (!file) return;
      // Don't re-upload the file we just pulled from Supabase ourselves
      if (input.dataset.fromSupabase === "1") { input.dataset.fromSupabase = ""; return; }
      pushFileToSupabase(slot, file);
    });
  });
}

// ── On auth ready: load whatever's already in Supabase for everyone ──
document.addEventListener("epss-auth-ready", async () => {
  attachAdminUploadSync();

  for (const slot of Object.keys(FILE_SLOTS)) {
    const input = document.getElementById(FILE_SLOTS[slot].inputId);
    if (input) input.dataset.fromSupabase = "1"; // mark so the change-sync above skips re-uploading
    await pullFileFromSupabase(slot);
  }
});
