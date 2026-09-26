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
    var commands = window.NexusPlmCommands;
    var toolbar = window.NexusPlmToolbar;

    var client = new clientModule.Client();
    var editorType = null;
    //: Attached once. Button ids never change, and attaching twice runs a command twice.
    var handlersAttached = false;
    //: The tray heartbeat, kept alive for as long as this panel is.
    var beat = null;
    //: What the tab and the panel are both drawn from, so they cannot disagree.
    var context = { signedIn: false, user: null, state: null };
    //: Handlers are attached once; button ids never change, so re-sending the tab does not
    //: need them re-attached, and attaching twice would run a command twice per click.
    var handlersAttached = false;

    function el(id) { return document.getElementById(id); }

    function draw(state, user, signedIn, trouble) {
        var t = trouble || {};
        context = { signedIn: signedIn, user: user, state: state };
        el("panel-title").textContent = rules.title(editorType);
        el("panel-headline").textContent = rules.headline(state, {
            hasDocument: true,
            signedIn: signedIn,
            hasSecret: t.hasSecret === false ? false : clientModule.hasSecret(),
            unreachable: t.unreachable,
            refused: t.refused
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

    /**
     * Run one command, from the tab or from the panel.
     *
     * Both go through here so that a command cannot behave differently depending on which one was
     * pressed - the bug the single command table exists to prevent.
     */
    function run(commandId) {
        var command = commands.byId(commandId);
        if (!command) { return; }

        // Never act on a command the rules say is not available. The tab greys it and the panel
        // disables it, but a stale tab is one editor repaint away and the service would then
        // refuse in a way the user cannot tell from a bug.
        if (!commands.isEnabled(command, context)) {
            say(commands.disabledBecause(command, context));
            return;
        }

        if (command.id === "navigator") { togglePanel(); return; }
        if (command.id === "help") { openHelp(); return; }

        var body = commands.bodyFor(command, context);
        if (body === null) { reportConnection(); return; }

        client.command(command.endpoint, body).then(function (answer) {
            if (answer && answer.unreachable) { say(answer.error, "error"); return; }
            if (answer && answer.success === false && answer.error) { say(answer.error, "warning"); }
            // Whatever it did, the document's state may have moved underneath us.
            refresh();
        });
    }

    /**
     * Every command, grouped the way the Word ribbon groups them.
     *
     * They live here rather than on a ribbon tab because Desktop Editors 9.4.0's editors
     * implement no menu-extension method at all - measured: AddToolbarMenuItem,
     * UpdateToolbarMenuItem and the context-menu methods appear nowhere in its sdk-all.js, so the
     * call is accepted by the plugin runtime and then goes nowhere. The tab payload is still
     * built and still sent, so a newer build lights it up with no further work; this is what a
     * user has today.
     */
    function drawCommands() {
        var host = el("panel-buttons");
        host.textContent = "";

        commands.GROUPS.forEach(function (group) {
            var members = commands.inGroup(group);
            if (!members.length) { return; }

            var heading = document.createElement("h2");
            heading.className = "nexus-group";
            heading.textContent = group;
            host.appendChild(heading);

            members.forEach(function (command) {
                var why = commands.disabledBecause(command, context);
                var button = document.createElement("button");
                button.className = "nexus-button";
                button.textContent = command.label;
                button.disabled = !commands.isEnabled(command, context);
                // A greyed button that does not say why is a dead end the user pokes at.
                button.title = why || command.label;
                // Each button carries what it does, rather than a handler working out which one
                // was pressed. The LibreOffice panel learned that the hard way.
                button.addEventListener("click", function () { run(command.id); });
                host.appendChild(button);
            });
        });
    }

    /** Say something through the tray's own toast, so it looks like it does in every host. */
    function say(message, severity) {
        if (!message) { return; }
        try { client.notify(message, severity || "info"); }
        catch (e) { window.console && window.console.log(message); }
    }

    /** Connection Status has no dialog of its own: it reports what we can see from here. */
    function reportConnection() {
        client.health().then(function (answer) {
            if (answer && answer.unreachable) { say(answer.error, "error"); }
            else if (context.signedIn) { say("Connected to Nexus PLM, signed in as " + context.user + ".", "success"); }
            else { say("Connected to Nexus PLM. Not signed in.", "info"); }
        });
    }

    function togglePanel() {
        // The Navigator IS this panel. Pressing it from the tab while the panel is open is a
        // no-op rather than a close: closing would take the tab's click handling with it.
    }

    function openHelp() {
        try { window.Asc.plugin.executeMethod("OpenLink", ["https://nexusplm.help/"]); }
        catch (e) { window.console && window.console.log("Nexus PLM: could not open help."); }
    }


    function refresh() {
        if (!clientModule.hasSecret()) { draw(null, null, false, { hasSecret: false }); return; }

        client.me().then(function (me) {
            // Three different failures, said three different ways. Reporting all of them as
            // "sign in" was wrong: only one of them is fixed by signing in.
            if (me && me.unreachable) { draw(null, null, false, { unreachable: true }); return; }
            if (me && me.refused) { draw(null, null, false, { refused: true }); return; }

            var signedIn = !!(me && me.success && me.username);
            if (!signedIn) { draw(null, null, false, {}); return; }

            // Which file this is comes from the editor, not from us; until that is wired the
            // panel draws the signed-in, document-unknown state, which is a real state.
            client.state({}).then(function (state) {
                draw(state, me.username, true, {});
            });
        });
    }

    /**
     * Put the Nexus PLM tab in the ribbon, and keep its buttons in step.
     *
     * This lives in the PANEL, not in a background variation, and that is the whole point: a
     * background variation is a one-shot action - it runs, is torn down, and its heartbeat stops.
     * The tab's registration survives that, but there is no frame left to receive a click, so
     * every button did nothing. This frame stays alive while the panel is open.
     *
     * Handlers are attached BEFORE the tab is sent, the order ONLYOFFICE's own plugins-ui.js uses.
     */
    function drawToolbar() {
        if (!window.Asc || !window.Asc.plugin) { return; }
        try {
            if (!handlersAttached) {
                commands.COMMANDS.forEach(function (command) {
                    window.Asc.plugin.attachToolbarMenuClickEvent(
                        toolbar.buttonId(command.id),
                        function () { run(command.id); });
                });
                handlersAttached = true;
            }
            window.Asc.plugin.executeMethod(
                "AddToolbarMenuItem", [toolbar.tab(window.Asc.plugin.guid, context)]);
        } catch (e) {
            // Put it on screen. There is no console to read in a plugin frame, and a tab that
            // silently fails to draw is the exact failure that cost this feature an evening.
            var headline = el("panel-headline");
            if (headline) { headline.textContent = "Tab error: " + (e && e.message ? e.message : e); }
            window.console && window.console.log("Nexus PLM: the tab could not be drawn - " + e.message);
        }
    }

    function start(type) {
        editorType = type;
        // Register with the tray, the way every other host does, so Nexus lists this one as
        // connected and can tell when it goes away.
        client.connect();
        beat = window.setInterval(function () { client.heartbeat(); }, clientModule.HEARTBEAT_MS);
        window.addEventListener("unload", function () {
            if (beat) { window.clearInterval(beat); beat = null; }
            client.disconnect();
        });
        refresh();
    }

    window.NexusPlmUi = {
        start: start,
        documentReady: refresh
    };

})(window, document);
