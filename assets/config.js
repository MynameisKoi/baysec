/**
 * BaySec Unified Configuration & Environment Routing
 * Decouples Production (Vercel) from Testing/Staging (GitHub Pages / Localhost)
 */
window.IS_PROD = window.location.hostname.includes("vercel.app") || window.location.hostname === "baysec.sfbu.edu";
window.FORCE_SIM = new URLSearchParams(window.location.search).get("sim") === "1";
window.SIM_MODE = !window.IS_PROD || window.FORCE_SIM;

var IS_PROD = window.IS_PROD;
var FORCE_SIM = window.FORCE_SIM;
var SIM_MODE = window.SIM_MODE;

var APPS_SCRIPT_URL = window.APPS_SCRIPT_URL || "https://script.google.com/macros/s/AKfycbzQO2sZyXWsUtTgpHzbGy5MLTyV-0-kQlnqddTG3SjwDcOMKsjrA7p5kwBjc_tJ7do_Cg/exec";
window.APPS_SCRIPT_URL = APPS_SCRIPT_URL;

// Synchronously initialize staging banner visibility based on SIM_MODE
function initSimBanner() {
    const simBanner = document.getElementById("simBanner");
    if (simBanner) {
        if (SIM_MODE) {
            simBanner.style.display = "flex";
        } else {
            simBanner.style.display = "none";
        }
    }
}

if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", initSimBanner);
} else {
    initSimBanner();
}
