/*
 * The Navigator panel's DOM: it puts what `lib/panel.js` decided on screen, and does nothing else.
 *
 * No rule about what to show lives here. If a question can be answered without an editor and
 * without a browser — which rows, which buttons are live, what to say when PLM has never seen the
 * document — it belongs in `lib/panel.js`, where the tests reach it. This file exists because a
 * DOM cannot be tested that way, and it is deliberately dull.
 *
 * The panel does not draw the tab and does not run commands itself: both belong to the runner
 * (`runner.js`), which the resident background half also uses, so a command pressed here does
 * exactly what the same command does on the tab.
 */

(function (window, document) {
    "use strict";

    var rules = window.NexusPlmPanel;
    var clientModule = window.NexusPlmClient;

    var runner = null;
    var editorType = null;
    var beat = null;

    function el(id) { return document.getElementById(id); }

    function draw(context) {
        var trouble = context.trouble || {};
        var doc = context.doc || { shell: false, path: null };

        el("panel-title").textContent = rules.title(editorType);
        el("panel-headline").textContent = rules.headline(context.state, {
            hasDocument: !doc.shell || !!doc.path,
            signedIn: context.signedIn,
            hasSecret: trouble.hasSecret === false ? false : clientModule.hasSecret(),
            unreachable: trouble.unreachable,
            refused: trouble.refused
        });

        var rows = el("panel-rows");
        rows.textContent = "";
        rules.rowsFor(context.state).forEach(function (pair) {
            var dt = document.createElement("dt");
            dt.textContent = pair[0];
            var dd = document.createElement("dd");
            dd.textContent = pair[1];
            rows.appendChild(dt);
            rows.appendChild(dd);
        });

        var allowed = rules.enabledButtons(context.state, context.user);
        var buttons = el("panel-buttons");
        buttons.textContent = "";

        rules.BUTTONS.forEach(function (b) {
            var button = document.createElement("button");
            button.textContent = b[0];
            button.className = "nexus-button";
            button.disabled = allowed.indexOf(b[1]) === -1;
            // Each button carries what it does, rather than the handler working out which button
            // was pressed. The LibreOffice panel learned that one the hard way.
            button.addEventListener("click", function () { runner.runId(b[1]); });
            buttons.appendChild(button);
        });
    }

    function start(type) {
        editorType = type;
        runner = window.NexusPlmRunner.create({ editorType: type, onState: draw });

        // Register with the tray, the way every other host does, so Nexus lists this one as
        // connected and can tell when it goes away.
        runner.client.connect();
        beat = window.setInterval(function () { runner.client.heartbeat(); }, clientModule.HEARTBEAT_MS);
        window.addEventListener("unload", function () {
            if (beat) { window.clearInterval(beat); beat = null; }
            runner.client.disconnect();
        });
        runner.refresh();
    }

    window.NexusPlmUi = {
        start: start,
        documentReady: function () { if (runner) { runner.refresh(); } }
    };

})(window, document);
