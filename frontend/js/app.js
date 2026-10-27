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
  if (!level) return "unassessed";
  return level.toLowerCase(); // "low" | "medium" | "high" | "unassessed"
}

/**
 * Return a hex color for the score bar fill.
 */
function scoreColor(level) {
  if (!level) return "#94a3b8";
  const map = { low: "#22c55e", medium: "#f59e0b", high: "#ef4444" };
  return map[level.toLowerCase()] || "#94a3b8";
}

/**
 * Build an HTML score bar element.
 */
function scoreBarHTML(score, level) {
  if (score === null || score === undefined || !level || level.toLowerCase() === "unassessed" || level.toLowerCase() === "none") {
    return `<span style="color:var(--muted); font-size:0.85rem; font-weight:600;">—</span>`;
  }
  const color = scoreColor(level);
  return `
    <div class="score-bar-wrap">
      <div class="score-bar">
        <div class="score-bar-fill" style="width:${score}%; background:${color};"></div>
      </div>
      <span class="score-label" style="color:${color};">${score}%</span>
    </div>`;
}

/**
 * Build a risk badge span.
 */
function badgeHTML(level) {
  const isAr = typeof i18n !== "undefined" && i18n.lang() === "ar";
  if (!level || level.toLowerCase() === "unassessed" || level.toLowerCase() === "none") {
    return `<span class="badge" style="background:#f1f5f9; color:#64748b; border:1px solid #cbd5e1; font-weight:600;">${isAr ? "⚪ لم يقيّم بعد" : "⚪ Unassessed"}</span>`;
  }
  const lvlLower = level.toLowerCase();
  let label = level;
  if (isAr) {
    if (lvlLower === "low") label = "🟢 منخفض";
    else if (lvlLower === "medium") label = "🟡 متوسط";
    else if (lvlLower === "high") label = "🔴 مرتفع";
  }
  return `<span class="badge badge-${levelClass(level)}">${label}</span>`;
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

/**
 * Generate intelligent, personalized wellness & recovery recommendations
 * tailored dynamically to the exact diagnostic outcome (Risk Level + Specific SHAP root causes).
 */
function getPersonalizedWellnessTips(diagnosisData) {
  if (!diagnosisData) return [];
  const isAr = typeof i18n !== "undefined" && i18n.lang() === "ar";
  const lvl = (diagnosisData.level || "Low").toLowerCase();
  const reasons = diagnosisData.top_reasons || [];
  const tips = [];

  if (lvl === "low") {
    return isAr ? [
      "✅ وضعك ممتاز ومستقر: حافظ على نمط العمل المتوازن وساعات الدوام الحالية.",
      "🌱 ضع حدوداً واضحة بين وقت العمل وحياتك الشخصية لتفادي تراكم الضغوط مستقبلاً.",
      "📚 استثمر هذا الاستقرار في تطوير مهاراتك أو مشاركة زملائك في خلق بيئة إيجابية.",
      "🏃 واظب على ممارسة الرياضة والأنشطة الترفيهية لتجديد طاقتك الذهنية والبدنية."
    ] : [
      "✅ Great balance: Your current workload metrics are healthy and sustainable.",
      "🌱 Protect your boundaries: Continue enforcing clear limits between work hours and personal life.",
      "📚 Growth opportunity: Use this mental clarity to invest in learning or mentoring team members.",
      "🏃 Maintain physical activity: Regular movement and hobbies help preserve cognitive endurance."
    ];
  }

  // If High or Medium risk: match tips to the exact root-cause drivers!
  const hasReason = (keyword) => reasons.some(r => r.toLowerCase().includes(keyword.toLowerCase()));

  // 1. Overtime hours driver
  if (hasReason("overtime")) {
    tips.push(isAr
      ? "⏱️ تخفيض الساعات الإضافية فوراً: ساعات الأوفرتايم تلتهم وقت استشفاء الجسم. اتفق مع مديرك على سقف أقصى لا يتجاوز 4 ساعات أسبوعياً."
      : "⏱️ Cap weekly overtime immediately: Extended overtime is depleting your biological recovery time. Agree with your lead on a strict overtime limit."
    );
  }

  // 2. High daily working hours driver
  if (hasReason("daily") || hasReason("hours")) {
    tips.push(isAr
      ? "🛑 ضبط ساعات الدوام اليومي: تجاوز 9 ساعات يومياً يرفع الكورتيزول ويخفض الإنتاجية. التزم بإنهاء الدوام في موعده وطبق تقنية البومودورو (25د تركيز + 5د راحة)."
      : "🛑 Bound daily shift hours: Working beyond standard hours spikes cortisol and diminishes returns. Hard-stop at shift end and take 5-min micro-breaks."
    );
  }

  // 3. High task load / multitasking driver
  if (hasReason("task")) {
    tips.push(isAr
      ? "📋 إعادة جدولة وتفويض المهام: عبء المهام المتراكمة يسبب تشتتاً ذهنياً. ركز على أهم 3 أولويات يومياً وفوض المهام الروتينية لزملائك."
      : "📋 Triage & delegate tasks: Context switching between excessive tasks causes cognitive fatigue. Pick top 3 daily priorities and delegate secondary items."
    );
  }

  // 4. Insufficient vacation days driver
  if (hasReason("vacation")) {
    tips.push(isAr
      ? "🏖️ جدولة إجازة سنوية عاجلة: رصيد إجازاتك المستهلك قليل جداً. احجز إجازة متصلة من 3 إلى 5 أيام لإعادة شحن طاقتك الذهنية والنفسية."
      : "🏖️ Schedule accrued annual leave: Your vacation utilization is critically low. Block 3-5 consecutive days off for complete mental detachment."
    );
  }

  // 5. Negative sentiment / psychological stress driver
  if (hasReason("sentiment") || hasReason("negative") || (diagnosisData.sentiment_risk && diagnosisData.sentiment_risk >= 0.65)) {
    tips.push(isAr
      ? "🗣️ دعم نفسي وتواصل شفاف: مؤشرات مشاعرك تشير لضغط نفسي وإحباط. احجز جلسة مكاشفة مع مسؤول الموارد البشرية (HR) أو مستشار نفسي لتفريغ الضغوط."
      : "🗣️ Transparent communication & mental support: High emotional strain detected. Schedule a safe check-in with HR or a wellness counselor."
    );
  }

  // 6. Absences or delays driver (symptom of disengagement / chronic fatigue)
  if (hasReason("absence") || hasReason("late")) {
    tips.push(isAr
      ? "😴 استعادة إيقاع النوم والراحة: تكرار التأخير والغياب مؤشر على الإرهاق التراكمي. احرص على 7-9 ساعات نوم منتظمة وناقش خيار الدوام المرن."
      : "😴 Sleep hygiene & fatigue recovery: Chronic tardiness/absences signal physical burnout. Prioritize 7-8 hours of uninterrupted sleep and explore flexible hours."
    );
  }

  // If High risk: prepend an urgent emergency action banner
  if (lvl === "high") {
    tips.unshift(isAr
      ? "🚨 تدخّل عاجل: أنت في منطقة الخطر الحرج للاحتراق الوظيفي. اطلب اجتماعاً فورياً لتعديل خطة العمل وتخفيف المسؤوليات قبل الوصول للانهيار البدني."
      : "🚨 Critical intervention needed: You are in the acute burnout zone. Request an urgent workload rebalancing meeting before physical exhaustion peaks."
    );
  } else if (tips.length < 3) {
    tips.push(isAr
      ? "🧘 روتين تفريغ التوتر: خصص 15 دقيقة يومياً للمشي الخفيف أو تمارين التنفس العميق لفصل ذهنك عن ضغوط العمل."
      : "🧘 Stress decompression routine: Dedicate 15 minutes daily to brisk walking or mindfulness to disconnect from work pressure."
    );
    tips.push(isAr
      ? "📵 فصل إشعارات العمل: أوقف تنبيهات البريد وتطبيقات العمل بمجرد مغادرة المكتب لاستعادة راحة البال."
      : "📵 Digital boundary: Mute corporate emails and chat notifications after shift hours to protect cognitive rest."
    );
  }

  return tips.slice(0, 5); // Return top targeted tips
}

// Run on load
document.addEventListener("DOMContentLoaded", setActiveNav);
