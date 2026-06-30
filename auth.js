// ════════════════════════════════════════════════════════════════
// auth.js — Supabase authentication + role gating
// Loaded BEFORE script.js / mos.js / etc. Blocks the app behind a
// login screen, then exposes window.isAdmin / window.APP_USER and
// hides admin-only UI for non-admins.
// ════════════════════════════════════════════════════════════════

// ── 1) FILL THESE IN from Supabase Dashboard → Project Settings → API ──
const SUPABASE_URL      = "https://qcccdwossgrotjiosurg.supabase.co";
const SUPABASE_ANON_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InFjY2Nkd29zc2dyb3RqaW9zdXJnIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODI3OTAxOTgsImV4cCI6MjA5ODM2NjE5OH0.pmDkCdQxAsh2f11MEe8MgdDy1vpTw6X6-R-gMiAEsq0";

const supabaseClient = window.supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY);
window.supabaseClient = supabaseClient; // used by storage-sync.js

window.APP_USER = null;   // { id, email, role }
window.isAdmin  = false;

// ── 2) BUILD THE LOGIN OVERLAY (injected, no HTML edits needed) ──
function injectAuthOverlay() {
  const el = document.createElement("div");
  el.id = "auth-overlay";
  el.innerHTML = `
    <div id="auth-box">
      <img src="epss-logo.png" alt="" id="auth-logo" />
      <h2>EPSS Stock-Multiple</h2>
      <p class="auth-sub">Sign in to continue</p>
      <form id="auth-form">
        <input type="email" id="auth-email" placeholder="Email" autocomplete="username" required />
        <input type="password" id="auth-password" placeholder="Password" autocomplete="current-password" required />
        <button type="submit" id="auth-submit">Sign In</button>
      </form>
      <div id="auth-error"></div>
      <div id="auth-loading" style="display:none">Checking session…</div>
    </div>
  `;
  document.body.appendChild(el);

  const style = document.createElement("style");
  style.textContent = `
    #auth-overlay {
      position: fixed; inset: 0; z-index: 9999;
      display: flex; align-items: center; justify-content: center;
      background: var(--bg, #0b1620);
    }
    #auth-box {
      width: 320px; max-width: 90vw; text-align: center;
      background: var(--card-bg, #0d1b2a); border: 1px solid var(--border, #1e2e3d);
      border-radius: 12px; padding: 2rem 1.6rem;
    }
    #auth-logo { height: 56px; margin: 0 auto 0.8rem; display: block; }
    #auth-box h2 { margin: 0 0 0.2rem; color: var(--text, #cdd9e5); font-size: 1.1rem; }
    .auth-sub { margin: 0 0 1.2rem; color: var(--muted, #7a97b0); font-size: 0.85rem; }
    #auth-form { display: flex; flex-direction: column; gap: 0.6rem; }
    #auth-form input {
      padding: 10px 12px; border-radius: 7px; border: 1.5px solid var(--border, #1e2e3d);
      background: var(--surface2, #0f2030); color: var(--text, #cdd9e5); font-size: 0.9rem;
    }
    #auth-form input:focus { outline: none; border-color: var(--blue, #3a8fd4); }
    #auth-submit {
      margin-top: 0.4rem; padding: 10px; border-radius: 7px; border: none;
      background: var(--blue, #3a8fd4); color: #fff; font-weight: 600; cursor: pointer;
    }
    #auth-submit:disabled { opacity: 0.6; cursor: not-allowed; }
    #auth-error { color: var(--red, #f85149); font-size: 0.8rem; margin-top: 0.8rem; min-height: 1em; }
    #auth-loading { color: var(--muted, #7a97b0); font-size: 0.85rem; margin-top: 1rem; }
  `;
  document.head.appendChild(style);
}

function showAuthOverlay() {
  const el = document.getElementById("auth-overlay");
  if (el) el.style.display = "flex";
  document.documentElement.style.overflow = "hidden";
}
function hideAuthOverlay() {
  const el = document.getElementById("auth-overlay");
  if (el) el.style.display = "none";
  document.documentElement.style.overflow = "";
}

// ── 3) ROLE-GATED UI ──
function applyRoleToUI() {
  const adminSection = document.getElementById("admin-upload-section");
  if (adminSection) adminSection.style.display = window.isAdmin ? "" : "none";

  let badge = document.getElementById("user-badge");
  if (!badge) {
    badge = document.createElement("div");
    badge.id = "user-badge";
    badge.style.cssText = "padding:0.6rem 0.9rem;font-size:0.75rem;color:var(--muted,#7a97b0);display:flex;align-items:center;justify-content:space-between;gap:0.5rem;";
    const sidebar = document.getElementById("sidebar");
    if (sidebar) sidebar.appendChild(badge);
  }
  if (window.APP_USER) {
    badge.innerHTML = `
      <span>${window.isAdmin ? "🛡️ Admin" : "👁️ Viewer"} · ${escapeHtml(window.APP_USER.email)}</span>
      <button id="logout-btn" style="background:none;border:1px solid var(--border,#1e2e3d);color:var(--muted,#7a97b0);border-radius:6px;padding:3px 8px;cursor:pointer;font-size:0.7rem">Sign out</button>
    `;
    document.getElementById("logout-btn").addEventListener("click", () => supabaseClient.auth.signOut());
  }
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, c => ({ "&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;" }[c]));
}

// ── 4) SESSION / PROFILE RESOLUTION ──
async function loadProfileAndUnlock(session) {
  const { data: profile, error } = await supabaseClient
    .from("profiles")
    .select("role,email")
    .eq("id", session.user.id)
    .single();

  if (error) {
    console.error("Could not load profile:", error);
    document.getElementById("auth-error").textContent = "Could not load your profile. Contact an admin.";
    return;
  }

  window.APP_USER = { id: session.user.id, email: profile.email || session.user.email, role: profile.role };
  window.isAdmin  = profile.role === "admin";

  hideAuthOverlay();
  applyRoleToUI();

  // Tell the rest of the app auth is ready (storage-sync.js listens for this)
  document.dispatchEvent(new CustomEvent("epss-auth-ready", { detail: window.APP_USER }));
}

async function initAuth() {
  injectAuthOverlay();
  showAuthOverlay();

  document.getElementById("auth-form").addEventListener("submit", async (e) => {
    e.preventDefault();
    const email    = document.getElementById("auth-email").value.trim();
    const password = document.getElementById("auth-password").value;
    const btn      = document.getElementById("auth-submit");
    const errEl    = document.getElementById("auth-error");
    errEl.textContent = "";
    btn.disabled = true;
    btn.textContent = "Signing in…";

    const { data, error } = await supabaseClient.auth.signInWithPassword({ email, password });

    btn.disabled = false;
    btn.textContent = "Sign In";

    if (error) {
      errEl.textContent = error.message;
      return;
    }
    await loadProfileAndUnlock(data.session);
  });

  // Restore existing session (page refresh)
  const { data: { session } } = await supabaseClient.auth.getSession();
  if (session) await loadProfileAndUnlock(session);

  supabaseClient.auth.onAuthStateChange((event, session) => {
    if (event === "SIGNED_OUT") {
      window.APP_USER = null;
      window.isAdmin  = false;
      location.reload(); // simplest way to fully reset in-memory app state
    }
  });
}

document.addEventListener("DOMContentLoaded", initAuth);
