// app.js
// ─────────────────────────────────────────────────────────
// Shared JS utilities used across all pages.
// ─────────────────────────────────────────────────────────

const API = CONFIG.API_BASE_URL;

// ── Helpers ───────────────────────────────────────────────

/**
 * Escape HTML to prevent Cross-Site Scripting (XSS).
 */
function escapeHTML(str) {
  if (str === null || str === undefined) return "";
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

/**
 * Return CSS class + badge class based on risk level string.
 */
function levelClass(level) {
  return level.toLowerCase(); // "low" | "medium" | "high"
}

/**
 * Return a hex color for the score bar fill.
 */
function scoreColor(level) {
  const map = { low: "#22c55e", medium: "#f59e0b", high: "#ef4444" };
  return map[level.toLowerCase()] || "#94a3b8";
}

/**
 * Build an HTML score bar element.
 */
function scoreBarHTML(score, level) {
  const color = scoreColor(level);
  return `
    <div class="score-bar-wrap">
      <div class="score-bar">
        <div class="score-bar-fill" style="width:${score}%; background:${color};"></div>
      </div>
      <span class="score-label" style="color:${color};">${score}</span>
    </div>`;
}

/**
 * Build a risk badge span.
 */
function badgeHTML(level) {
  return `<span class="badge badge-${levelClass(level)}">${level}</span>`;
}

/**
 * Mark active nav link.
 */
function setActiveNav() {
  const page = location.pathname.split("/").pop();
  document.querySelectorAll("nav a").forEach(a => {
    const href = a.getAttribute("href").split("/").pop();
    a.classList.toggle("active", href === page);
  });
}

// Run on load
document.addEventListener("DOMContentLoaded", setActiveNav);
