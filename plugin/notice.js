/*
 * The notice window's only job: show the message it was opened with.
 *
 * The message arrives through the plugin's own storage - `nexusplm.notice.v1`, written by
 * editor.js just before this window is opened. It travelled in the URL fragment first, and that
 * was measured on Desktop Editors 9.4.0: a plugin window whose url carries a `#` never loads its
 * page at all. It shows the editor's "Loading" spinner for ever, logs nothing, and the only way
 * out is to kill the editor - so the window meant to say "the tray is not running" became a
 * worse failure than the silence it was written to fix. The Navigator's `index.html`, with no
 * fragment, has always loaded.
 *
 * Nothing is decided here; the frame that opened this window owns its OK button (ONLYOFFICE
 * reports a plugin window's buttons to the OPENER's Asc.plugin.button), so this page has no
 * handlers of its own beyond the two the SDK expects to find.
 */

(function (window, document) {
    "use strict";

    var KEY = "nexusplm.notice.v1";

    window.Asc = window.Asc || {};
    window.Asc.plugin = window.Asc.plugin || {};

    /** What the opener left in storage, or a line saying the message did not survive the trip. */
    function payload() {
        try {
            var raw = window.localStorage && window.localStorage.getItem(KEY);
            var parsed = raw ? JSON.parse(raw) : null;
            if (parsed && typeof parsed.message === "string") { return parsed; }
        } catch (e) { /* fall through */ }
        // Never blank: a window with nothing in it is the failure this page exists to avoid.
        return {
            message: "Nexus PLM could not read what it was going to tell you. "
                   + "Is the Nexus PLM tray application running?",
            severity: "warning"
        };
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

    // Drawn at load as well as from init: the SDK calls init once, and a page that waited for it
    // and did not get it would show the empty box this page exists to avoid.
    draw();

})(window, document);
