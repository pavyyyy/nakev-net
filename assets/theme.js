(function () {
  "use strict";

  var cookieName = "siteTheme";

  function readCookie(name) {
    return document.cookie.split(";").map(function (part) {
      return part.trim();
    }).reduce(function (found, part) {
      if (found) return found;
      var prefix = name + "=";
      return part.indexOf(prefix) === 0 ? decodeURIComponent(part.slice(prefix.length)) : "";
    }, "");
  }

  function writeCookie(name, value) {
    var maxAge = 60 * 60 * 24 * 365;
    document.cookie = name + "=" + encodeURIComponent(value) + "; max-age=" + maxAge + "; path=/; SameSite=Lax";
  }

  function preferredTheme() {
    var saved = readCookie(cookieName);
    if (saved === "light" || saved === "dark") return saved;
    return window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light";
  }

  function applyTheme(theme, save) {
    var next = theme === "dark" ? "dark" : "light";
    document.documentElement.setAttribute("data-theme", next);
    document.documentElement.classList.toggle("dark", next === "dark");
    if (save) writeCookie(cookieName, next);

    document.querySelectorAll("[data-theme-toggle], #darkModeToggle").forEach(function (control) {
      if (control.type === "checkbox") {
        control.checked = next === "dark";
      } else {
        control.setAttribute("aria-pressed", String(next === "dark"));
        control.textContent = next === "dark" ? "Light" : "Dark";
      }
    });
  }

  applyTheme(preferredTheme(), false);

  document.addEventListener("DOMContentLoaded", function () {
    applyTheme(document.documentElement.getAttribute("data-theme") || preferredTheme(), false);
    document.querySelectorAll("[data-theme-toggle], #darkModeToggle").forEach(function (control) {
      control.addEventListener("change", function () {
        applyTheme(control.checked ? "dark" : "light", true);
      });
      control.addEventListener("click", function () {
        if (control.type !== "checkbox") {
          applyTheme(document.documentElement.getAttribute("data-theme") === "dark" ? "light" : "dark", true);
        }
      });
    });
  });
})();
