/*
 * The notice window's only job: show the message it was opened with.
 *
 * The message arrives in the URL fragment as JSON - `{ message, severity }` - put there by
 * editor.js. Nothing is decided here; the frame that opened this window owns its OK button
 * (ONLYOFFICE reports a plugin window's buttons to the OPENER's Asc.plugin.button), so this
 * page has no handlers of its own beyond the two the SDK expects to find.
 */

(function (window, document) {
    "use strict";

    window.Asc = window.Asc || {};
    window.Asc.plugin = window.Asc.plugin || {};

    /** What the fragment carries, or a message that says the page was opened without one. */
    function payload() {
        try {
            var raw = String(window.location.hash || "").replace(/^#/, "");
            var parsed = raw ? JSON.parse(decodeURIComponent(raw)) : null;
            if (parsed && typeof parsed.message === "string") { return parsed; }
        } catch (e) { /* fall through */ }
        return { message: "Nexus PLM has nothing to report.", severity: "info" };
    }

    function draw() {
        var p = payload();
        var box = document.getElementById("np-notice");
        if (!box) { return; }
        box.textContent = p.message;
        if (p.severity === "warning" || p.severity === "error") {
            box.className = "np-notice np-notice-" + p.severity;
        }
    }

    window.Asc.plugin.init = function () { draw(); };
    window.Asc.plugin.button = function () {};
    window.Asc.plugin.onExternalMouseUp = function () {};

    // Draw at load as well: a window's init is only ever called once by the SDK, and drawing
    // twice costs nothing.
    draw();

})(window, document);
