// auth.js — Client-side Session and Role Guard
const AUTH = {
  getUser() {
    const raw = localStorage.getItem("wp_user");
    return raw ? JSON.parse(raw) : null;
  },

  setUser(user) {
    localStorage.setItem("wp_user", JSON.stringify(user));
  },

  logout() {
    localStorage.removeItem("wp_user");
    window.location.href = "login.html";
  },

  requireAdmin() {
    const user = this.getUser();
    if (!user || user.role !== "admin") {
      alert(i18n ? (i18n.lang() === "ar" ? "غير مصرح لك بالدخول، هذه الصفحة للأدمن فقط" : "Unauthorized: Admin access required") : "Admin access required");
      window.location.href = "login.html";
    }
  },

  requireLogin() {
    const user = this.getUser();
    if (!user) {
      window.location.href = "login.html";
    }
    return user;
  }
};
