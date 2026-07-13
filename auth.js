// ════════════════════════════════════════════════════════════════
// auth.js — Supabase authentication + role gating
// Loaded BEFORE script.js / mos.js / etc. Shows a full landing +
// login page to signed-out visitors (sidebar & app stay hidden
// behind it), then exposes window.isAdmin / window.APP_USER and
// hides admin-only UI for non-admins once signed in.
// ════════════════════════════════════════════════════════════════

// ── 1) FILL THESE IN from Supabase Dashboard → Project Settings → API ──
const SUPABASE_URL      = "https://qcccdwossgrotjiosurg.supabase.co";
const SUPABASE_ANON_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InFjY2Nkd29zc2dyb3RqaW9zdXJnIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODI3OTAxOTgsImV4cCI6MjA5ODM2NjE5OH0.pmDkCdQxAsh2f11MEe8MgdDy1vpTw6X6-R-gMiAEsq0";

const supabaseClient = window.supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY);
window.supabaseClient = supabaseClient; // used by storage-sync.js

window.APP_USER = null;   // { id, email, role }
window.isAdmin  = false;

// ── 2) BUILD THE LANDING + LOGIN OVERLAY (injected, no HTML edits needed) ──
const AUTH_FEATURES = [
  { icon: "📊", title: "Dashboard Overview",        desc: "Aggregated inventory metrics across every plant and material group, updated the moment new data lands." },
  { icon: "🚚", title: "Stock in Transit",           desc: "See exactly what's moving between plants and how long it's been on the road." },
  { icon: "⏰", title: "Expiry Watchlist",           desc: "Catch batches heading toward expiry early enough to actually do something about it." },
  { icon: "🔬", title: "Quality Inspection",         desc: "Track material held in QC so nothing sits in limbo without anyone noticing." },
  { icon: "🏢", title: "Branch Comparison",          desc: "Compare stock positions across branches side by side to spot imbalances fast." },
  { icon: "🎯", title: "Stock Concentration",        desc: "Identify where inventory is overly concentrated and where it's dangerously thin." },
  { icon: "🧯", title: "Overstock & Expiry Risk",    desc: "Model redistribution plans that move at-risk stock before it becomes a write-off." },
  { icon: "🗂️", title: "National Stock & MOS",       desc: "One row per material, network-wide — stock on hand and Months of Stock at a glance." },
];

function injectAuthOverlay() {
  const el = document.createElement("div");
  el.id = "auth-overlay";
  el.innerHTML = `
    <nav id="auth-nav">
      <div class="auth-nav-inner">
        <div class="auth-nav-brand">
          <img src="epss-logo.png" alt="" />
          <span>EPSS Stock-Multiple</span>
        </div>
        <div class="auth-nav-links">
          <a href="#auth-features">Features</a>
          <a href="#auth-about">About</a>
        </div>
        <button type="button" class="auth-nav-cta" id="auth-nav-login-btn">Log In</button>
      </div>
    </nav>

    <div id="auth-scroll">
      <section id="auth-hero">
        <div class="auth-hero-grid">
          <div class="auth-hero-text">
            <span class="auth-hero-eyebrow">Demand & Inventory Management Directorate</span>
            <h1>Pharmaceutical Inventory Management</h1>
            <p>Track stock across every plant, catch expiry risk before it becomes loss, and always know who's responsible for what — all from one dashboard.</p>
            <div class="auth-hero-actions">
              <button type="button" class="auth-btn-primary" id="auth-hero-signin-btn">→ Sign In</button>
              <a href="#auth-features" class="auth-btn-secondary">ⓘ Learn More</a>
            </div>
            <div class="auth-hero-pills">
              <span class="auth-pill">⏰ Expiry Tracking</span>
              <span class="auth-pill">🏢 Branch Comparison</span>
              <span class="auth-pill">🗂️ National MOS</span>
            </div>
          </div>

          <div class="auth-login-card" id="auth-login-card">
            <div class="auth-login-header">
              <span class="auth-login-icon">🔒</span>
              <h2>Login to EPSS Stock-Multiple Track</h2>
            </div>

            <form id="auth-form">
              <label class="auth-field-label" for="auth-email">Email Address</label>
              <input type="email" id="auth-email" placeholder="you@epss.gov.et" autocomplete="username" required />

              <label class="auth-field-label" for="auth-password">Password</label>
              <input type="password" id="auth-password" placeholder="••••••••" autocomplete="current-password" required />

              <label class="auth-remember-row">
                <input type="checkbox" id="auth-remember" checked />
                <span>Remember Me</span>
              </label>

              <button type="submit" id="auth-submit">→ Login</button>
              <button type="button" id="auth-forgot-btn">Forgot Your Password?</button>
            </form>

            <div id="auth-error"></div>
            <div id="auth-loading">Checking session…</div>
          </div>
        </div>
      </section>

      <section id="auth-features">
        <div class="auth-section-inner">
          <div class="auth-section-head">
            <span class="auth-hero-eyebrow">What's inside</span>
            <h2>Everything your team needs to stay ahead of stock-outs and expiry</h2>
          </div>
          <div class="auth-feature-grid">
            ${AUTH_FEATURES.map(f => `
              <div class="auth-feature-card">
                <div class="auth-feature-icon">${f.icon}</div>
                <div class="auth-feature-title">${f.title}</div>
                <div class="auth-feature-desc">${f.desc}</div>
              </div>
            `).join("")}
          </div>
        </div>
      </section>

      <section id="auth-about">
        <div class="auth-section-inner auth-about-inner">
          <div>
            <span class="auth-hero-eyebrow">About</span>
            <h2>Built for EPSS inventory teams</h2>
            <p>EPSS Stock-Multiple pulls inventory, transit, and AMC data into a single network-wide view — so decisions about redistribution, expiry risk, and ownership are based on what's actually on the shelf, not what's in someone's spreadsheet.</p>
          </div>
          <button type="button" class="auth-btn-primary" id="auth-about-signin-btn">→ Sign In to Get Started</button>
        </div>
      </section>

      <footer id="auth-footer">© <span id="auth-year"></span> EPSS Stock-Multiple · Inventory Management</footer>
    </div>
  `;
  document.body.appendChild(el);

  const yearEl = el.querySelector("#auth-year");
  if (yearEl) yearEl.textContent = new Date().getFullYear();

  const style = document.createElement("style");
  style.textContent = `
    #auth-overlay {
      position: fixed; inset: 0; z-index: 9999;
      display: flex; flex-direction: column;
      background: var(--bg, #07090d);
      color: var(--text, #dce8f5);
      font-family: 'Inter', system-ui, sans-serif;
    }
    #auth-scroll { flex: 1; overflow-y: auto; }

    /* ── Top nav ── */
    #auth-nav {
      border-bottom: 1px solid var(--border, #1f2e44);
      background: var(--surface, #0e1420);
      flex-shrink: 0;
    }
    .auth-nav-inner {
      max-width: 1180px; margin: 0 auto;
      display: flex; align-items: center; gap: 1.2rem;
      padding: 0.8rem 1.5rem;
    }
    .auth-nav-brand { display: flex; align-items: center; gap: 0.6rem; font-weight: 700; font-size: 0.95rem; }
    .auth-nav-brand img { height: 30px; width: auto; display: block; }
    .auth-nav-links { display: flex; gap: 1.4rem; margin-left: auto; }
    .auth-nav-links a { color: var(--muted, #7a9ab8); text-decoration: none; font-size: 0.85rem; font-weight: 500; transition: color 0.15s; }
    .auth-nav-links a:hover { color: var(--text, #dce8f5); }
    .auth-nav-cta {
      background: var(--blue, #3d94e0); color: #fff; border: none;
      border-radius: 999px; padding: 0.5rem 1.1rem; font-size: 0.82rem;
      font-weight: 600; cursor: pointer; transition: opacity 0.15s; font-family: inherit;
    }
    .auth-nav-cta:hover { opacity: 0.88; }

    /* ── Hero ── */
    #auth-hero {
      background:
        radial-gradient(circle at 15% 20%, var(--blue-glow, rgba(61,148,224,0.22)) 0%, transparent 45%),
        radial-gradient(circle at 85% 80%, rgba(148,113,214,0.18) 0%, transparent 45%),
        var(--bg, #07090d);
      padding: 3.2rem 1.5rem;
    }
    .auth-hero-grid {
      max-width: 1180px; margin: 0 auto;
      display: grid; grid-template-columns: 1.1fr 0.9fr; gap: 3rem; align-items: center;
    }
    .auth-hero-eyebrow {
      display: inline-block; font-size: 0.72rem; font-weight: 700; letter-spacing: 0.04em;
      color: var(--blue, #3d94e0); background: var(--blue-glow, rgba(61,148,224,0.16));
      border: 1px solid var(--blue-soft, #1b4a70); border-radius: 999px; padding: 0.3rem 0.75rem;
      margin-bottom: 1rem;
    }
    .auth-hero-text h1 {
      font-size: 2.5rem; line-height: 1.15; margin: 0 0 1rem; font-weight: 800;
      font-family: 'Plus Jakarta Sans', 'Inter', sans-serif;
    }
    .auth-hero-text p { font-size: 1rem; color: var(--muted, #7a9ab8); max-width: 46ch; margin: 0 0 1.6rem; line-height: 1.55; }
    .auth-hero-actions { display: flex; gap: 0.8rem; flex-wrap: wrap; margin-bottom: 1.6rem; }
    .auth-btn-primary, .auth-btn-secondary {
      display: inline-flex; align-items: center; gap: 0.4rem;
      padding: 0.75rem 1.4rem; border-radius: 999px; font-size: 0.88rem; font-weight: 700;
      cursor: pointer; text-decoration: none; font-family: inherit; border: none;
      transition: opacity 0.15s, transform 0.1s;
    }
    .auth-btn-primary { background: var(--blue, #3d94e0); color: #fff; }
    .auth-btn-primary:hover { opacity: 0.9; }
    .auth-btn-secondary { background: var(--surface2, #141c2b); color: var(--text, #dce8f5); border: 1px solid var(--border, #1f2e44); }
    .auth-btn-secondary:hover { border-color: var(--blue, #3d94e0); color: var(--blue, #3d94e0); }
    .auth-hero-pills { display: flex; flex-wrap: wrap; gap: 0.55rem; }
    .auth-pill {
      font-size: 0.74rem; font-weight: 600; color: var(--muted, #7a9ab8);
      background: var(--surface2, #141c2b); border: 1px solid var(--border, #1f2e44);
      border-radius: 999px; padding: 0.4rem 0.85rem;
    }

    /* ── Login card ── */
    .auth-login-card {
      background: var(--surface, #0e1420); border: 1px solid var(--border, #1f2e44);
      border-radius: var(--radius-lg, 14px); padding: 1.8rem 1.7rem;
      box-shadow: 0 20px 50px rgba(0,0,0,0.45);
    }
    .auth-login-header { display: flex; align-items: center; justify-content: center; gap: 0.5rem; margin-bottom: 1.4rem; }
    .auth-login-icon { font-size: 1.1rem; }
    .auth-login-header h2 { margin: 0; font-size: 1.05rem; font-weight: 700; }
    #auth-form { display: flex; flex-direction: column; }
    .auth-field-label { font-size: 0.74rem; font-weight: 600; color: var(--muted, #7a9ab8); margin: 0.7rem 0 0.35rem; }
    .auth-field-label:first-child { margin-top: 0; }
    #auth-form input[type="email"], #auth-form input[type="password"] {
      padding: 10px 12px; border-radius: var(--radius-md, 10px); border: 1.5px solid var(--border, #1f2e44);
      background: var(--surface2, #141c2b); color: var(--text, #dce8f5); font-size: 0.9rem; font-family: inherit;
      width: 100%; box-sizing: border-box;
    }
    #auth-form input:focus { outline: none; border-color: var(--blue, #3d94e0); box-shadow: 0 0 0 3px var(--blue-glow, rgba(61,148,224,0.22)); }
    .auth-remember-row {
      display: flex; align-items: center; gap: 0.45rem; margin: 0.9rem 0 0.2rem;
      font-size: 0.78rem; color: var(--muted, #7a9ab8); cursor: pointer;
    }
    .auth-remember-row input { accent-color: var(--blue, #3d94e0); }
    #auth-submit {
      margin-top: 1rem; padding: 11px; border-radius: var(--radius-md, 10px); border: none;
      background: var(--blue, #3d94e0); color: #fff; font-weight: 700; font-size: 0.9rem;
      cursor: pointer; font-family: inherit; transition: opacity 0.15s;
    }
    #auth-submit:hover { opacity: 0.9; }
    #auth-submit:disabled { opacity: 0.6; cursor: not-allowed; }
    #auth-forgot-btn {
      background: none; border: none; color: var(--blue, #3d94e0); font-size: 0.78rem;
      cursor: pointer; margin-top: 0.8rem; text-decoration: underline; font-family: inherit; padding: 0;
    }
    #auth-error { color: var(--red, #e04545); font-size: 0.8rem; margin-top: 0.9rem; min-height: 1em; text-align: center; }
    #auth-loading { display: none; color: var(--muted, #7a9ab8); font-size: 0.82rem; margin-top: 1rem; text-align: center; }
    #auth-loading.show { display: block; }

    /* ── Features section ── */
    #auth-features, #auth-about { padding: 3.2rem 1.5rem; }
    #auth-features { background: var(--surface, #0e1420); border-top: 1px solid var(--border, #1f2e44); }
    .auth-section-inner { max-width: 1180px; margin: 0 auto; }
    .auth-section-head { max-width: 60ch; margin-bottom: 2rem; }
    .auth-section-head h2 { font-size: 1.5rem; margin: 0.4rem 0 0; font-family: 'Plus Jakarta Sans', sans-serif; }
    .auth-feature-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 1rem; }
    .auth-feature-card {
      background: var(--surface2, #141c2b); border: 1px solid var(--border, #1f2e44);
      border-radius: var(--radius-md, 10px); padding: 1.1rem;
      transition: border-color 0.15s, transform 0.15s;
    }
    .auth-feature-card:hover { border-color: var(--blue, #3d94e0); transform: translateY(-2px); }
    .auth-feature-icon { font-size: 1.4rem; margin-bottom: 0.5rem; }
    .auth-feature-title { font-weight: 700; font-size: 0.88rem; margin-bottom: 0.35rem; }
    .auth-feature-desc { font-size: 0.76rem; color: var(--muted, #7a9ab8); line-height: 1.45; }

    /* ── About ── */
    .auth-about-inner {
      display: flex; align-items: center; justify-content: space-between; gap: 2rem; flex-wrap: wrap;
    }
    .auth-about-inner h2 { font-size: 1.4rem; margin: 0.4rem 0 0.7rem; font-family: 'Plus Jakarta Sans', sans-serif; }
    .auth-about-inner p { color: var(--muted, #7a9ab8); max-width: 56ch; line-height: 1.6; margin: 0; font-size: 0.9rem; }

    #auth-footer {
      text-align: center; padding: 1.4rem; font-size: 0.74rem; color: var(--dim, #4a6275);
      border-top: 1px solid var(--border, #1f2e44); background: var(--surface, #0e1420);
    }

    @media (max-width: 900px) {
      .auth-hero-grid { grid-template-columns: 1fr; }
      .auth-hero-text h1 { font-size: 2rem; }
      .auth-feature-grid { grid-template-columns: repeat(2, 1fr); }
    }
    @media (max-width: 640px) {
      .auth-nav-links { display: none; }
      .auth-feature-grid { grid-template-columns: 1fr; }
      .auth-about-inner { flex-direction: column; align-items: flex-start; }
    }
  `;
  document.head.appendChild(style);

  // Nav / hero CTAs scroll to + focus the login card
  const focusLogin = () => {
    document.getElementById("auth-login-card").scrollIntoView({ behavior: "smooth", block: "center" });
    setTimeout(() => document.getElementById("auth-email").focus(), 350);
  };
  document.getElementById("auth-nav-login-btn").addEventListener("click", focusLogin);
  document.getElementById("auth-hero-signin-btn").addEventListener("click", focusLogin);
  document.getElementById("auth-about-signin-btn").addEventListener("click", focusLogin);

  // Smooth-scroll the in-page nav anchors
  el.querySelectorAll('a[href^="#"]').forEach(a => {
    a.addEventListener("click", (e) => {
      const target = document.querySelector(a.getAttribute("href"));
      if (!target) return;
      e.preventDefault();
      target.scrollIntoView({ behavior: "smooth", block: "start" });
    });
  });

  // Forgot password
  document.getElementById("auth-forgot-btn").addEventListener("click", async () => {
    const errEl = document.getElementById("auth-error");
    const email = document.getElementById("auth-email").value.trim();
    if (!email) {
      errEl.style.color = "var(--red, #e04545)";
      errEl.textContent = "Enter your email above first, then click \"Forgot Your Password?\"";
      return;
    }
    errEl.style.color = "var(--muted, #7a9ab8)";
    errEl.textContent = "Sending reset link…";
    const { error } = await supabaseClient.auth.resetPasswordForEmail(email);
    if (error) {
      errEl.style.color = "var(--red, #e04545)";
      errEl.textContent = error.message;
      return;
    }
    errEl.style.color = "var(--green, #30a85f)";
    errEl.textContent = "Reset link sent — check your inbox.";
  });
}

function showAuthOverlay() {
  const el = document.getElementById("auth-overlay");
  if (el) el.style.display = "flex";
  document.documentElement.style.overflow = "hidden";
  document.getElementById("auth-scroll").scrollTop = 0;
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
    const errEl = document.getElementById("auth-error");
    if (errEl) { errEl.style.color = "var(--red, #e04545)"; errEl.textContent = "Could not load your profile. Contact an admin."; }
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

  const loadingEl = document.getElementById("auth-loading");
  loadingEl.classList.add("show");

  document.getElementById("auth-form").addEventListener("submit", async (e) => {
    e.preventDefault();
    const email    = document.getElementById("auth-email").value.trim();
    const password = document.getElementById("auth-password").value;
    const btn      = document.getElementById("auth-submit");
    const errEl    = document.getElementById("auth-error");
    errEl.style.color = "var(--red, #e04545)";
    errEl.textContent = "";
    btn.disabled = true;
    btn.textContent = "Signing in…";

    const { data, error } = await supabaseClient.auth.signInWithPassword({ email, password });

    btn.disabled = false;
    btn.textContent = "→ Login";

    if (error) {
      errEl.textContent = error.message;
      return;
    }
    await loadProfileAndUnlock(data.session);
  });

  // Restore existing session (page refresh)
  const { data: { session } } = await supabaseClient.auth.getSession();
  loadingEl.classList.remove("show");
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
