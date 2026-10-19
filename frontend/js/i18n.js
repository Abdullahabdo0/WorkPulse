// i18n.js — Bilingual (Arabic / English) translations
// Usage: i18n.t("key") returns the string in the current language.

const TRANSLATIONS = {
  en: {
    // Nav
    nav_dashboard:   "Dashboard",
    nav_employees:   "Employees",
    nav_assessment:  "New Assessment",
    nav_portal:      "Employee Portal",
    lang_toggle:     "العربية",

    // Auth
    auth_portal_title:    "WorkPulse Access",
    auth_portal_subtitle: "AI Workplace Burnout Management System",
    auth_notice:          "Authorized access only. Employee accounts are created and managed exclusively by HR Administration.",
    auth_tab_login:       "Login",
    auth_username:        "Username",
    auth_password:        "Password",
    auth_fullname:        "Full Name",
    auth_btn_login:       "Sign In",

    // Dashboard
    dash_title:           "Executive Analytics Dashboard",
    dash_subtitle:        "Live enterprise workload intelligence & predictive burnout monitoring",
    dash_total:           "Total Employees",
    dash_total_sub:       "Active workforce",
    dash_low:             "Healthy & Balanced",
    dash_medium:          "Moderate Workload",
    dash_high:            "Critical Burnout Risk 🔴",
    dash_kpi_avg_hours:   "Daily Hours Avg",
    dash_kpi_avg_ot:      "Weekly Overtime Avg",
    dash_kpi_avg_tasks:   "Avg Active Tasks",
    dash_kpi_avg_score:   "Workforce Risk Index",
    dash_chart_dist:      "Risk Level Distribution",
    dash_chart_bar:       "Employees by Risk Category",
    dash_chart_benchmark: "Workload Benchmark by Risk Group",
    dash_watchlist_title: "🚨 Critical Attention Watchlist",
    dash_watchlist_sub:   "Employees with highest burnout probability requiring immediate workload rebalancing",
    dash_col_emp:         "Employee",
    dash_col_score:       "Risk Score",
    dash_col_ot:          "Overtime",
    dash_col_action:      "Action",
    dash_btn_inspect:     "Inspect",
    dash_api_error:       "Could not connect to the API. Make sure the backend is running: uvicorn backend.main:app --reload",

    // Employees
    emp_title:            "👥 Employee Workforce Directory",
    emp_subtitle:         "Comprehensive burnout risk assessment, monitoring and profile management",
    emp_btn_add:          "+ Add New Employee",
    emp_search:           "Search by Name, Username or Employee ID...",
    emp_filter_all:       "All Risk Levels",
    emp_filter_low:       "🟢 Low Risk",
    emp_filter_med:       "🟡 Medium Risk",
    emp_filter_high:      "🔴 High Risk",
    emp_col_id:           "Employee ID",
    emp_col_name:         "Full Name",
    emp_col_username:     "Username",
    emp_col_hours:        "Daily Hours",
    emp_col_ot:           "Overtime / Wk",
    emp_col_abs:          "Absences",
    emp_col_tasks:        "Tasks",
    emp_col_score:        "Burnout Score",
    emp_col_level:        "Risk Level",
    emp_col_actions:      "Actions",
    emp_btn_edit:         "Edit",
    emp_btn_delete:       "Delete",
    emp_loading:          "Loading workforce directory...",
    emp_no_data:          "No employees found matching criteria.",
    emp_showing:          "Showing",
    emp_of:               "of",
    emp_employees:        "employees",
    emp_confirm_delete:   "Are you sure you want to permanently delete this employee? This will wipe the account and all assessments.",
    emp_success_add:      "Employee created and AI assessed successfully!",
    emp_success_update:   "Employee profile updated and re-assessed!",
    emp_success_delete:   "Employee deleted successfully.",
    emp_modal_add_title:  "Add New Employee & Create Account",
    emp_modal_creds_hint: "Enter employee account credentials. They will use this username and password to log in.",
    emp_modal_edit_title: "Edit Employee Profile & Workload",
    emp_modal_sec_creds:  "1. Account Credentials (Login Data)",
    emp_modal_sec_work:   "2. Workload & Operational Metrics",
    emp_modal_new_pwd:    "New Password (leave empty to keep current)",
    emp_modal_btn_cancel: "Cancel",
    emp_modal_btn_save:   "Save Changes",
    emp_modal_btn_create: "Create Employee",
    emp_api_error:        "Could not load employee data. Make sure backend is running.",

    // Assessment
    assess_title:         "New Burnout Risk Assessment",
    assess_hours_label:   "Average Daily Hours *",
    assess_hours_hint:    "Range: 4 – 16 hours/day",
    assess_ot_label:      "Overtime Hours / Week *",
    assess_ot_hint:       "Range: 0 – 30 hours/week",
    assess_abs_label:     "Absence Days (last month) *",
    assess_abs_hint:      "Range: 0 – 15 days",
    assess_late_label:    "Late Arrivals (last month) *",
    assess_late_hint:     "Range: 0 – 20 times",
    assess_tasks_label:   "Tasks Assigned *",
    assess_tasks_hint:    "Range: 1 – 60 tasks",
    assess_vac_label:     "Vacation Days (last year) *",
    assess_vac_hint:      "Range: 0 – 30 days",
    assess_fb_label:      "Employee Feedback / Sentiment * (English only)",
    assess_fb_placeholder:"How does the employee feel about their workload? (English only)",
    assess_fb_hint:       "Write in English — the AI sentiment model supports English only.",
    assess_quickfill:     "Quick fill:",
    assess_high_ex:       "High Risk Example",
    assess_med_ex:        "Medium Risk Example",
    assess_low_ex:        "Low Risk Example",
    assess_btn:           "Analyze Risk",
    assess_btn_loading:   "Analyzing...",
    assess_result_sub:    "Burnout Risk Score (0 – 100)",
    assess_workload_lbl:  "Workload Probability",
    assess_sentiment_lbl: "Sentiment Risk",
    assess_reasons_title: "Top Contributing Factors",
    assess_level_low:     "Low Risk",
    assess_level_med:     "Medium Risk",
    assess_level_high:    "High Risk",
    assess_api_err:       "Make sure the backend is running (uvicorn backend.main:app --reload)",

    // Employee Portal
    portal_title:         "Employee Self-Assessment Portal",
    portal_subtitle:      "Enter your name to check your burnout risk assessment",
    portal_name_search_label:  "Your Full Name *",
    portal_name_search_ph:     "e.g. Ahmed Mohamed",
    portal_name_search_hint:   "Enter the name registered with HR.",
    portal_find_btn:      "Find My Record",
    portal_or:            "— or submit a new self-assessment —",
    portal_self_title:    "Self-Assessment Form",
    portal_self_sub:      "Fill in your work details honestly. Results are confidential.",
    portal_emp_name_label: "Your Name (optional)",
    portal_emp_name_ph:   "e.g. Ahmed Mohamed",
    portal_result_title:  "Your Risk Assessment Result",
    portal_result_note:   "This result will be reviewed by HR. It is a support tool — not a final decision.",
    portal_privacy:       "Your data is handled confidentially and used only for HR support purposes.",
    portal_notfound:      "No matching employee found. Please fill the form below manually.",
    portal_found:         "Record found! Your data has been pre-filled.",
    portal_multi_found:   "Multiple records found. Showing the first match.",
  },

  ar: {
    // Nav
    nav_dashboard:   "لوحة التحكم",
    nav_employees:   "الموظفين",
    nav_assessment:  "تقييم جديد",
    nav_portal:      "بوابة الموظف",
    lang_toggle:     "English",

    // Auth
    auth_portal_title:    "بوابة الدخول — WorkPulse",
    auth_portal_subtitle: "نظام الذكاء الاصطناعي لإدارة الإرهاق الوظيفي",
    auth_notice:          "الدخول مصرح به للموظفين والإدارة. يتم إنشاء وتعيين حسابات الموظفين عبر مسؤول الموارد البشرية (HR Admin) فقط.",
    auth_tab_login:       "تسجيل الدخول",
    auth_username:        "اسم المستخدم",
    auth_password:        "كلمة المرور",
    auth_fullname:        "الاسم بالكامل",
    auth_btn_login:       "دخول",

    // Dashboard
    dash_title:           "لوحة التحليلات التنفيذية المتقدمة",
    dash_subtitle:        "مراقبة حية ودقيقة لعبء العمل ومؤشرات الاحتراق الوظيفي بالذكاء الاصطناعي",
    dash_total:           "إجمالي الموظفين",
    dash_total_sub:       "القوى العاملة المسجلة",
    dash_low:             "بيئة صحية ومستقرة",
    dash_medium:          "عبء عمل متوسط",
    dash_high:            "خطر احتراق مرتفع 🔴",
    dash_kpi_avg_hours:   "متوسط ساعات العمل اليومية",
    dash_kpi_avg_ot:      "متوسط الساعات الإضافية أسبوعياً",
    dash_kpi_avg_tasks:   "متوسط المهام المسندة",
    dash_kpi_avg_score:   "مؤشر المخاطر العام للمؤسسة",
    dash_chart_dist:      "توزيع مستويات الخطر",
    dash_chart_bar:       "عدد الموظفين حسب فئة الخطر",
    dash_chart_benchmark: "مقارنة أعباء العمل ومؤشرات الإجهاد",
    dash_watchlist_title: "🚨 قائمة المتابعة الفورية (أعلى الموظفين عرضة للاحتراق)",
    dash_watchlist_sub:   "موظفون بحاجة لتدخل وتخفيف أعباء فوري من إدارة الموارد البشرية",
    dash_col_emp:         "الموظف",
    dash_col_score:       "مؤشر الخطر",
    dash_col_ot:          "الإضافي",
    dash_col_action:      "إجراء",
    dash_btn_inspect:     "فحص وتعديل",
    dash_api_error:       "تعذّر الاتصال بالـ API. تأكد من تشغيل الباك إند: uvicorn backend.main:app --reload",

    // Employees
    emp_title:            "👥 دليل وإدارة شؤون الموظفين",
    emp_subtitle:         "مراقبة مستمرة، تعديل بيانات، وتقييم ذكي لمستوى الاحتراق الوظيفي لكل موظف",
    emp_btn_add:          "+ إضافة موظف جديد",
    emp_search:           "ابحث بالاسم، اسم المستخدم أو الرقم الوظيفي...",
    emp_filter_all:       "جميع مستويات الخطر",
    emp_filter_low:       "🟢 خطر منخفض",
    emp_filter_med:       "🟡 خطر متوسط",
    emp_filter_high:      "🔴 خطر مرتفع",
    emp_col_id:           "الرقم الوظيفي",
    emp_col_name:         "الاسم بالكامل",
    emp_col_username:     "اسم المستخدم",
    emp_col_hours:        "ساعات/يوم",
    emp_col_ot:           "إضافي/أسبوع",
    emp_col_abs:          "الغياب",
    emp_col_tasks:        "المهام",
    emp_col_score:        "مؤشر الخطر",
    emp_col_level:        "المستوى",
    emp_col_actions:      "الإجراءات",
    emp_btn_edit:         "تعديل",
    emp_btn_delete:       "حذف",
    emp_loading:          "جارٍ تحميل بيانات الموظفين...",
    emp_no_data:          "لا توجد نتائج مطابقة لبحثك.",
    emp_showing:          "عرض",
    emp_of:               "من أصل",
    emp_employees:        "موظف",
    emp_confirm_delete:   "هل أنت متأكد من حذف هذا الموظف نهائياً؟ سيتم مسح حسابه وسجلاته بالكامل.",
    emp_success_add:      "تم إنشاء حساب الموظف وحساب مؤشر الخطر بالذكاء الاصطناعي بنجاح!",
    emp_success_update:   "تم تحديث بيانات الموظف وإعادة احتساب المؤشرات بنجاح!",
    emp_success_delete:   "تم حذف الموظف وجميع سجلاته بنجاح.",
    emp_modal_add_title:  "إضافة موظف جديد وإنشاء حسابه",
    emp_modal_creds_hint: "أدخل بيانات حساب الموظف. سيستخدم اسم المستخدم وكلمة المرور هذه لتسجيل الدخول.",
    emp_modal_edit_title: "تعديل بيانات وساعات عمل الموظف",
    emp_modal_sec_creds:  "1. بيانات الحساب وتسجيل الدخول",
    emp_modal_sec_work:   "2. مؤشرات ساعات العمل وسجل الأداء",
    emp_modal_new_pwd:    "كلمة المرور الجديدة (اتركها فارغة للإبقاء على الحالية)",
    emp_modal_btn_cancel: "إلغاء",
    emp_modal_btn_save:   "حفظ التعديلات",
    emp_modal_btn_create: "إنشاء وإضافة الموظف",
    emp_api_error:        "تعذّر تحميل البيانات. تأكد من تشغيل الباك إند.",

    // Assessment
    assess_title:         "تقييم جديد لخطر الإرهاق الوظيفي",
    assess_hours_label:   "متوسط ساعات العمل اليومية *",
    assess_hours_hint:    "المدى: 4 – 16 ساعة/يوم",
    assess_ot_label:      "ساعات الأوفرتايم الأسبوعية *",
    assess_ot_hint:       "المدى: 0 – 30 ساعة/أسبوع",
    assess_abs_label:     "أيام الغياب (الشهر الماضي) *",
    assess_abs_hint:      "المدى: 0 – 15 يوم",
    assess_late_label:    "مرات التأخير (الشهر الماضي) *",
    assess_late_hint:     "المدى: 0 – 20 مرة",
    assess_tasks_label:   "عدد المهام المسندة *",
    assess_tasks_hint:    "المدى: 1 – 60 مهمة",
    assess_vac_label:     "أيام الإجازة (السنة الماضية) *",
    assess_vac_hint:      "المدى: 0 – 30 يوم",
    assess_fb_label:      "ملاحظات الموظف * (بالإنجليزية فقط)",
    assess_fb_placeholder:"اكتب شعور الموظف تجاه ضغط العمل (بالإنجليزية فقط)",
    assess_fb_hint:       "اكتب بالإنجليزية — موديل تحليل المشاعر يدعم الإنجليزية فقط حالياً.",
    assess_quickfill:     "أمثلة سريعة:",
    assess_high_ex:       "مثال خطر مرتفع",
    assess_med_ex:        "مثال خطر متوسط",
    assess_low_ex:        "مثال خطر منخفض",
    assess_btn:           "تحليل الخطر",
    assess_btn_loading:   "جارٍ التحليل...",
    assess_result_sub:    "درجة خطر الإرهاق (0 – 100)",
    assess_workload_lbl:  "احتمالية الإرهاق (عبء العمل)",
    assess_sentiment_lbl: "خطر المشاعر",
    assess_reasons_title: "أبرز العوامل المؤثرة",
    assess_level_low:     "خطر منخفض",
    assess_level_med:     "خطر متوسط",
    assess_level_high:    "خطر مرتفع",
    assess_api_err:       "تأكد من تشغيل الباك إند (uvicorn backend.main:app --reload)",

    // Employee Portal
    portal_title:         "بوابة التقييم الذاتي للموظف",
    portal_subtitle:      "أدخل اسمك للاطلاع على مستوى خطر الإرهاق الوظيفي لديك",
    portal_name_search_label:  "اسمك الكامل *",
    portal_name_search_ph:     "مثال: أحمد محمد",
    portal_name_search_hint:   "أدخل الاسم المسجّل لدى الموارد البشرية.",
    portal_find_btn:      "البحث عن سجلي",
    portal_or:            "— أو أكمل التقييم الذاتي أدناه —",
    portal_self_title:    "نموذج التقييم الذاتي",
    portal_self_sub:      "أجب بصدق — النتائج سرية ولن تُشارَك إلا مع HR.",
    portal_emp_name_label: "اسمك (اختياري)",
    portal_emp_name_ph:   "مثال: أحمد محمد",
    portal_result_title:  "نتيجة تقييم الخطر الخاص بك",
    portal_result_note:   "هذه النتيجة ستُراجَع من HR. هي أداة دعم وليست قراراً نهائياً.",
    portal_privacy:       "بياناتك تُعالَج بسرية تامة وتُستخدم فقط لأغراض دعم الموارد البشرية.",
    portal_notfound:      "لم يُعثر على اسمك. يرجى ملء النموذج يدوياً أدناه.",
    portal_found:         "تم العثور على سجلك! تم تعبئة بياناتك تلقائياً.",
    portal_multi_found:   "تم العثور على أكثر من سجل. يتم عرض أول نتيجة مطابقة.",
  }
};

// ── Language state ──────────────────────────────────────────
let currentLang = localStorage.getItem("wp_lang") || "en";

const i18n = {
  t(key) {
    return TRANSLATIONS[currentLang][key] || TRANSLATIONS["en"][key] || key;
  },
  lang() { return currentLang; },
  toggle() {
    currentLang = currentLang === "en" ? "ar" : "en";
    localStorage.setItem("wp_lang", currentLang);
    applyLanguage();
  },
  set(lang) {
    currentLang = lang;
    localStorage.setItem("wp_lang", lang);
    applyLanguage();
  }
};

/** Apply language to the document: set dir, update all [data-i18n] elements */
function applyLanguage() {
  const isAr = currentLang === "ar";
  document.documentElement.lang = currentLang;
  document.body.dir = isAr ? "rtl" : "ltr";

  // Update all elements with data-i18n attribute
  document.querySelectorAll("[data-i18n]").forEach(el => {
    const key = el.getAttribute("data-i18n");
    const val = i18n.t(key);
    if (el.tagName === "INPUT" || el.tagName === "TEXTAREA") {
      el.placeholder = val;
    } else {
      el.textContent = val;
    }
  });

  // Update placeholder-only elements
  document.querySelectorAll("[data-i18n-ph]").forEach(el => {
    el.placeholder = i18n.t(el.getAttribute("data-i18n-ph"));
  });

  // Update the lang toggle button text
  document.querySelectorAll(".lang-btn").forEach(btn => {
    btn.textContent = i18n.t("lang_toggle");
  });

  // Dispatch event so page-specific JS can react
  document.dispatchEvent(new CustomEvent("langchange", { detail: currentLang }));
}

// Apply on first load
document.addEventListener("DOMContentLoaded", applyLanguage);
