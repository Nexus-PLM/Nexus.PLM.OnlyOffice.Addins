/*
 * The panel's DOM: it puts what `lib/panel.js` decided on screen, and does nothing else.
 *
 * No rule about what to show lives here. If a question can be answered without an editor and
 * without a browser — which rows, which buttons are live, what to say when PLM has never seen the
 * document — it belongs in `lib/panel.js`, where the tests reach it. This file exists because a
 * DOM cannot be tested that way, and it is deliberately dull.
 */

(function (window, document) {
    "use strict";

    var rules = window.NexusPlmPanel;
    var clientModule = window.NexusPlmClient;

    var client = new clientModule.Client();
    var editorType = null;

    function el(id) { return document.getElementById(id); }

    function draw(state, user, signedIn) {
        el("panel-title").textContent = rules.title(editorType);
        el("panel-headline").textContent = rules.headline(state, {
            hasDocument: true,
            signedIn: signedIn
        });

        var rows = el("panel-rows");
        rows.textContent = "";
        rules.rowsFor(state).forEach(function (pair) {
            var dt = document.createElement("dt");
            dt.textContent = pair[0];
            var dd = document.createElement("dd");
            dd.textContent = pair[1];
            rows.appendChild(dt);
            rows.appendChild(dd);
        });

        var allowed = rules.enabledButtons(state, user);
        var buttons = el("panel-buttons");
        buttons.textContent = "";
        rules.BUTTONS.forEach(function (b) {
            var button = document.createElement("button");
            button.textContent = b[0];
            button.className = "nexus-button";
            button.disabled = allowed.indexOf(b[1]) === -1;
            // Each button carries what it does, rather than the handler working out which button
            // was pressed. The LibreOffice panel learned that one the hard way.
            button.addEventListener("click", function () { run(b[1]); });
            buttons.appendChild(button);
        });
    }

    function run(command) {
        // The commands themselves are the next slice; see docs/COMMANDS.md for the set and the
        // order they are being built in. Wiring them before the service can be reached from a
        // browser frame at all would be building on an unmeasured assumption.
        window.console && window.console.log("Nexus PLM: " + command + " is not wired up yet.");
    }

    function refresh() {
        client.me().then(function (me) {
            var signedIn = !!(me && me.success && me.username);
            if (!signedIn) { draw(null, null, false); return; }
            // Which file this is comes from the editor, not from us; until that is wired the
            // panel draws the signed-in, document-unknown state, which is a real state.
            client.state({}).then(function (state) {
                draw(state, me.username, true);
            });
        });
    }

    function start(type) {
        editorType = type;
        refresh();
    }

    window.NexusPlmUi = {
        start: start,
        documentReady: refresh
    };

})(window, document);
