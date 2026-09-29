/*
 * Background email job progress toast.
 *
 * Usage (from a page):
 *     ApolloEmailJob.start('send', payload, recipients)
 *       payload    -> body for POST /email-jobs (recipients, subject, body, signature)
 *       recipients -> [{ email, name }] used for display
 *
 * The job id is kept in localStorage, so the toast resumes on any page
 * that includes this script. When a job finishes, the page receives an
 * `apollo:email-job-done` event with the final job data.
 */
(function () {
    "use strict";

    var STORAGE_KEY = "apollo-email-job";
    var POLL_MS = 2000;

    var state = null;      // { id, action, names: {email: name}, minimized, finishedNotified }
    var pollTimer = null;
    var toast = null;
    var lastStatus = {};   // email -> status, to animate only fresh "sent" ticks

    function load() {
        try {
            return JSON.parse(localStorage.getItem(STORAGE_KEY) || "null");
        } catch (e) {
            return null;
        }
    }

    function save() {
        try {
            if (state) localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
            else localStorage.removeItem(STORAGE_KEY);
        } catch (e) {}
    }

    function esc(value) {
        return String(value == null ? "" : value)
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;");
    }

    function initials(name, email) {
        var source = (name || email || "?").trim();
        var parts = source.split(/[\s._@-]+/).filter(Boolean);
        return ((parts[0] || "?").charAt(0) + (parts[1] || "").charAt(0)) || "?";
    }

    function ensureToast() {
        if (toast) return toast;

        toast = document.createElement("div");
        toast.className = "ej-toast";
        toast.setAttribute("role", "status");
        toast.setAttribute("aria-live", "polite");
        toast.innerHTML =
            '<div class="ej-head">' +
                '<div class="ej-head-icon"><i class="fas fa-paper-plane"></i></div>' +
                '<div class="ej-head-text">' +
                    '<div class="ej-title">Preparing emails…</div>' +
                    '<div class="ej-subtitle">Starting</div>' +
                '</div>' +
                '<button type="button" class="ej-head-btn ej-min" title="Minimize" aria-label="Minimize"><i class="fas fa-chevron-down"></i></button>' +
                '<button type="button" class="ej-head-btn ej-close" title="Close" aria-label="Close"><i class="fas fa-times"></i></button>' +
            '</div>' +
            '<div class="ej-bar"><div class="ej-bar-fill"></div></div>' +
            '<ul class="ej-list"></ul>' +
            '<div class="ej-foot"><span class="ej-foot-text">You can keep working or open another page.</span></div>';

        document.body.appendChild(toast);

        toast.querySelector(".ej-min").addEventListener("click", function () {
            if (!state) return;
            state.minimized = !state.minimized;
            save();
            applyMinimized();
        });

        toast.querySelector(".ej-close").addEventListener("click", function () {
            if (state && !state.done) {
                // Closing while running only hides the list; the job keeps going.
                state.minimized = true;
                save();
                applyMinimized();
                return;
            }
            dismiss();
        });

        requestAnimationFrame(function () {
            toast.classList.add("ej-show");
        });

        return toast;
    }

    function applyMinimized() {
        if (!toast || !state) return;
        toast.classList.toggle("ej-minimized", Boolean(state.minimized));
        var icon = toast.querySelector(".ej-min i");
        if (icon) icon.className = state.minimized ? "fas fa-chevron-up" : "fas fa-chevron-down";
    }

    function dismiss() {
        stopPolling();
        state = null;
        save();
        if (toast) {
            toast.classList.remove("ej-show");
            var node = toast;
            toast = null;
            setTimeout(function () { node.remove(); }, 300);
        }
    }

    function statusHtml(item) {
        switch (item.status) {
            case "sent":
                return '<span class="ej-status ej-sent"><i class="fas fa-envelope-circle-check' +
                    (lastStatus[item.email] !== "sent" ? " ej-fly" : "") + '"></i> Sent</span>';
            case "queued":
                return '<span class="ej-status ej-queued"><i class="fas fa-clock"></i> Queued</span>';
            case "failed":
                return '<span class="ej-status ej-failed"><i class="fas fa-circle-xmark"></i> Failed</span>';
            case "sending":
                return '<span class="ej-status ej-sending"><span class="ej-spinner"></span> Sending</span>';
            default:
                return '<span class="ej-status ej-waiting"><i class="far fa-circle"></i> Waiting</span>';
        }
    }

    function render(data) {
        ensureToast();
        applyMinimized();

        var items = data.items || [];
        var counts = data.counts || {};
        var total = data.total || items.length || 0;
        var done = data.done || 0;
        var isQueue = data.action === "queue";
        var finished = data.status === "done";

        var title = toast.querySelector(".ej-title");
        var subtitle = toast.querySelector(".ej-subtitle");
        var headIcon = toast.querySelector(".ej-head-icon i");

        toast.classList.toggle("ej-state-done", finished && !counts.failed);
        toast.classList.toggle("ej-state-failed", finished && counts.failed > 0);

        if (finished) {
            title.textContent = counts.failed
                ? (isQueue ? "Queue finished with errors" : "Sending finished with errors")
                : (isQueue ? "Emails queued" : "All emails processed");
            headIcon.className = counts.failed ? "fas fa-triangle-exclamation" : "fas fa-check";
        } else {
            title.textContent = isQueue ? "Queueing emails…" : "Sending emails…";
            headIcon.className = "fas fa-paper-plane";
        }

        var parts = [done + " of " + total + " done"];
        if (counts.sent) parts.push(counts.sent + " sent");
        if (counts.queued) parts.push(counts.queued + " queued");
        if (counts.failed) parts.push(counts.failed + " failed");
        subtitle.textContent = parts.join(" · ");

        toast.querySelector(".ej-bar-fill").style.width =
            (total ? Math.round((done / total) * 100) : 0) + "%";

        var names = (state && state.names) || {};
        toast.querySelector(".ej-list").innerHTML = items.map(function (item) {
            var name = names[item.email] || item.email;
            var extra = "";
            if (item.status === "failed" && item.error) {
                extra = '<div class="ej-error" title="' + esc(item.error) + '">' + esc(item.error) + "</div>";
            } else if (item.status === "queued" && item.queued_for) {
                extra = '<div class="ej-note">Scheduled for ' + esc(item.queued_for) + " IST</div>";
            }
            return (
                '<li class="ej-item ej-' + esc(item.status) + '">' +
                    '<div class="ej-avatar">' + esc(initials(name, item.email)) + "</div>" +
                    '<div class="ej-item-text">' +
                        '<div class="ej-name">' + esc(name) + "</div>" +
                        '<div class="ej-email">' + esc(item.email) + "</div>" +
                        extra +
                    "</div>" +
                    statusHtml(item) +
                "</li>"
            );
        }).join("");

        items.forEach(function (item) {
            lastStatus[item.email] = item.status;
        });

        var foot = toast.querySelector(".ej-foot-text");
        if (finished) {
            foot.innerHTML = counts.failed
                ? 'Failed emails are listed above with the reason. <a href="/sent-emails">Open Sent Emails</a>'
                : '<a href="' + (isQueue ? "/queued-emails" : "/sent-emails") + '">' +
                  (isQueue ? "Open Queue" : "Open Sent Emails") + "</a>";
        } else {
            foot.textContent = "You can keep working or open another page.";
        }
    }

    function stopPolling() {
        if (pollTimer) {
            clearTimeout(pollTimer);
            pollTimer = null;
        }
    }

    function poll() {
        stopPolling();
        if (!state) return;

        fetch("/email-jobs/" + encodeURIComponent(state.id), { cache: "no-store" })
            .then(function (response) {
                if (response.status === 404) {
                    return { missing: true };
                }
                return response.json();
            })
            .then(function (data) {
                if (!state) return;

                if (data.missing) {
                    ensureToast();
                    toast.querySelector(".ej-title").textContent = "Progress no longer available";
                    toast.querySelector(".ej-subtitle").textContent =
                        "The server restarted. Check Sent Emails for the final status.";
                    state.done = true;
                    save();
                    return;
                }

                render(data);

                if (data.status === "done") {
                    state.done = true;
                    if (!state.finishedNotified) {
                        state.finishedNotified = true;
                        window.dispatchEvent(new CustomEvent("apollo:email-job-done", { detail: data }));
                    }
                    save();
                    return;
                }

                pollTimer = setTimeout(poll, POLL_MS);
            })
            .catch(function () {
                // Network hiccup: keep trying.
                pollTimer = setTimeout(poll, POLL_MS * 2);
            });
    }

    function start(action, payload, recipients) {
        var names = {};
        (recipients || []).forEach(function (r) {
            var email = String(r.email || "").trim().toLowerCase();
            if (email) names[email] = r.name || email;
        });

        state = { id: null, action: action, names: names, minimized: false, done: false };
        lastStatus = {};
        ensureToast();
        render({
            action: action,
            status: "running",
            total: Object.keys(names).length,
            done: 0,
            counts: {},
            items: Object.keys(names).map(function (email, index) {
                return { email: email, status: index === 0 ? "sending" : "waiting" };
            })
        });

        var body = Object.assign({}, payload, { action: action });

        return fetch("/email-jobs", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(body)
        })
            .then(function (response) {
                return response.json().then(function (data) {
                    if (!response.ok || !data.success) {
                        throw new Error(data.error || ("HTTP " + response.status));
                    }
                    return data;
                });
            })
            .then(function (data) {
                state.id = data.job_id;
                save();
                poll();
                return data;
            })
            .catch(function (error) {
                ensureToast();
                toast.classList.add("ej-state-failed");
                toast.querySelector(".ej-title").textContent = "Could not start sending";
                toast.querySelector(".ej-subtitle").textContent = error.message;
                state.done = true;
                save();
                throw error;
            });
    }

    function resume() {
        state = load();
        if (!state || !state.id) {
            state = null;
            return;
        }
        ensureToast();
        applyMinimized();
        poll();
    }

    window.ApolloEmailJob = {
        start: start,
        dismiss: dismiss,
        isRunning: function () {
            return Boolean(state && state.id && !state.done);
        }
    };

    if (document.readyState === "loading") {
        document.addEventListener("DOMContentLoaded", resume);
    } else {
        resume();
    }
})();
