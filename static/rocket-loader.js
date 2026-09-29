/*
 * Rocket page loader.
 *
 * Include right after <body> (not deferred) so the loader is visible while
 * the rest of the page loads:
 *     <script src="/static/rocket-loader.js"></script>
 *
 * - Hides once the page has fully loaded (minimum display time keeps it from flashing).
 * - Shows again on same-site link clicks and normal form submissions.
 * - window.showPageLoader(text) / window.hidePageLoader() for manual use.
 */
(function () {
    "use strict";

    // Apply the saved theme now, before the page renders, so dark mode
    // doesn't flash light first (each page's own theme script runs at the end).
    try {
        var savedTheme = localStorage.getItem("apollo-theme");
        if (savedTheme) document.documentElement.setAttribute("data-theme", savedTheme);
    } catch (e) {}

    var MIN_VISIBLE_MS = 700;
    var startedAt = Date.now();

    var ROCKET_SVG =
        '<svg viewBox="0 0 64 128" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">' +
            '<defs>' +
                '<linearGradient id="rlFin" x1="0" y1="0" x2="0" y2="1">' +
                    '<stop offset="0" stop-color="#3fb6e0"/><stop offset="1" stop-color="#1f5fa8"/>' +
                '</linearGradient>' +
                '<linearGradient id="rlNose" x1="0" y1="0" x2="1" y2="1">' +
                    '<stop offset="0" stop-color="#3fb6e0"/><stop offset="1" stop-color="#1f5fa8"/>' +
                '</linearGradient>' +
            '</defs>' +
            // side fins
            '<path d="M14 78 L2 104 L16 100 Z" fill="url(#rlFin)"/>' +
            '<path d="M50 78 L62 104 L48 100 Z" fill="url(#rlFin)"/>' +
            // body
            '<path d="M32 2 C50 18 52 52 50 100 L14 100 C12 52 14 18 32 2 Z" fill="#ffffff"/>' +
            // body shading
            '<path d="M32 2 C50 18 52 52 50 100 L38 100 C40 60 40 24 32 2 Z" fill="#d6f2f7"/>' +
            // nose cone
            '<path d="M32 2 C40 9 45 18 47 28 L17 28 C19 18 24 9 32 2 Z" fill="url(#rlNose)"/>' +
            // window
            '<circle cx="32" cy="50" r="10" fill="#0b2a5b" stroke="#3fb6e0" stroke-width="3"/>' +
            '<circle cx="28.5" cy="46.5" r="3" fill="#d6f2f7"/>' +
            // center fin
            '<path d="M30 82 L34 82 L35 104 L29 104 Z" fill="#1f5fa8"/>' +
            // nozzle
            '<rect x="20" y="100" width="24" height="8" rx="2" fill="#0b2a5b"/>' +
        '</svg>';

    function buildStars() {
        var html = "";
        for (var i = 0; i < 26; i++) {
            var left = Math.random() * 100;
            var duration = 0.6 + Math.random() * 1.2;
            var delay = Math.random() * 1.8;
            var height = 30 + Math.random() * 70;
            html +=
                '<span style="left:' + left.toFixed(2) + '%;' +
                'height:' + height.toFixed(0) + 'px;' +
                'animation-duration:' + duration.toFixed(2) + 's;' +
                'animation-delay:' + delay.toFixed(2) + 's;"></span>';
        }
        return html;
    }

    function createLoader() {
        var loader = document.createElement("div");
        loader.id = "rocketLoader";
        loader.className = "rocket-loader";
        loader.setAttribute("role", "status");
        loader.setAttribute("aria-live", "polite");
        loader.innerHTML =
            '<div class="rocket-loader-stars">' + buildStars() + '</div>' +
            '<div class="rocket-loader-stage">' +
                '<div class="rocket-loader-orbit"></div>' +
                '<div class="rocket-loader-planet"></div>' +
                '<div class="rocket-loader-rocket">' +
                    ROCKET_SVG +
                    '<div class="rocket-loader-flame"></div>' +
                    '<div class="rocket-loader-smoke"><span></span><span></span><span></span><span></span></div>' +
                '</div>' +
            '</div>' +
            '<div class="rocket-loader-title">Sales <span>Intelligence</span></div>' +
            '<div class="rocket-loader-subtitle" id="rocketLoaderText">Launching your workspace</div>' +
            '<div class="rocket-loader-progress"></div>';
        return loader;
    }

    var loader = document.getElementById("rocketLoader");
    if (!loader) {
        loader = createLoader();
        (document.body || document.documentElement).appendChild(loader);
    }
    document.documentElement.classList.add("rocket-loader-active");

    var hideTimer = null;

    function showPageLoader(text) {
        clearTimeout(hideTimer);
        var label = document.getElementById("rocketLoaderText");
        if (label) label.textContent = text || "Loading";
        startedAt = Date.now();
        loader.classList.remove("rocket-loader-hidden");
        document.documentElement.classList.add("rocket-loader-active");
    }

    function hidePageLoader() {
        var wait = Math.max(0, MIN_VISIBLE_MS - (Date.now() - startedAt));
        clearTimeout(hideTimer);
        hideTimer = setTimeout(function () {
            loader.classList.add("rocket-loader-hidden");
            document.documentElement.classList.remove("rocket-loader-active");
        }, wait);
    }

    window.showPageLoader = showPageLoader;
    window.hidePageLoader = hidePageLoader;

    if (document.readyState === "complete") {
        hidePageLoader();
    } else {
        window.addEventListener("load", hidePageLoader);
    }

    // Returning with the browser Back button can restore a page from cache
    // with the loader still showing.
    window.addEventListener("pageshow", function (event) {
        if (event.persisted) hidePageLoader();
    });

    function isModifiedClick(event) {
        return event.button !== 0 || event.metaKey || event.ctrlKey || event.shiftKey || event.altKey;
    }

    // Show on same-site navigation links.
    document.addEventListener("click", function (event) {
        if (event.defaultPrevented || isModifiedClick(event)) return;

        var link = event.target.closest && event.target.closest("a[href]");
        if (!link) return;
        if (link.target && link.target !== "_self") return;
        if (link.hasAttribute("download") || link.dataset.noLoader !== undefined) return;

        var href = link.getAttribute("href") || "";
        if (!href || href.charAt(0) === "#" || /^(javascript|mailto|tel):/i.test(href)) return;

        var url;
        try {
            url = new URL(link.href, window.location.href);
        } catch (e) {
            return;
        }
        if (url.origin !== window.location.origin) return;
        // Same page, only the hash changes.
        if (url.pathname === window.location.pathname && url.search === window.location.search && url.hash) return;

        // Let other click handlers run first; skip if one of them cancelled navigation.
        setTimeout(function () {
            if (!event.defaultPrevented) showPageLoader("Loading");
        }, 0);
    });

    // Show on normal (non-AJAX) form submissions.
    document.addEventListener("submit", function (event) {
        var form = event.target;
        if (!form || form.dataset.noLoader !== undefined) return;
        if (form.target && form.target !== "_self") return;

        setTimeout(function () {
            if (!event.defaultPrevented) showPageLoader("Fetching results");
        }, 0);
    });
})();
