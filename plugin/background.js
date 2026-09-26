/*
 * The Nexus PLM tab, and nothing else.
 *
 * This half of the plugin has no window. It puts the tab in the ribbon, keeps its buttons in step
 * with the signed-in user and the open document, and runs what the user presses.
 *
 * ── Why it is separate from the panel ────────────────────────────────────────────────────────
 * A visual variation is unloaded the moment its window is closed, and everything it registered
 * goes with it — measured on Desktop Editors 9.4.0: the tab appeared when the panel opened and
 * vanished the moment it was closed. A background variation keeps running for as long as the
 * document is open, so the tab stays. `Send` and `Encryption`, both shipped with Desktop Editors,
 * are built the same way.
 *
 * Every decision — which commands exist, when each may be pressed, what a greyed one should say —
 * is in `lib/commands.js` and shared with the panel, so the tab and the panel cannot disagree.
 */

(function (window) {
    "use strict";

    var commands = window.NexusPlmCommands;
    var toolbar = window.NexusPlmToolbar;
    var clientModule = window.NexusPlmClient;

    var client = new clientModule.Client();

    /** What the tab is drawn from. The panel keeps its own copy of the same three things. */
    var context = { signedIn: false, user: null, state: null };

    /** Attached once: button ids never change, and attaching twice runs a command twice. */
    var handlersAttached = false;

    /** The heartbeat timer, so it can be stopped when the document closes. */
    var beat = null;

    window.Asc = window.Asc || {};
    window.Asc.plugin = window.Asc.plugin || {};

    window.Asc.plugin.init = function () {
        // Register with the tray first, so Nexus lists this host as connected the way it lists
        // Word and FreeCAD. Without it the plugin works but is invisible - which is exactly how
        // it looked: a tab full of commands and nothing in the tray.
        client.connect();
        beat = window.setInterval(function () { client.heartbeat(); }, clientModule.HEARTBEAT_MS);

        draw();
        refresh();
    };

    /** Tell the tray we have gone, rather than leaving it to time us out. */
    window.Asc.plugin.onExternalMouseUp = function () {};
    window.addEventListener("unload", function () {
        if (beat) { window.clearInterval(beat); beat = null; }
        client.disconnect();
    });

    /** The document changed under us, so what the commands may do may have changed too. */
    window.Asc.plugin.onDocumentContentReady = function () {
        refresh();
    };

    window.Asc.plugin.button = function () {
        // A background variation has no buttons of its own; nothing to do.
    };

    /**
     * Put the tab in the ribbon.
     *
     * The whole tab is re-sent on every state change rather than patched: AddToolbarMenuItem
     * replaces the tab for this guid, and one call that is always right beats a diff that is
     * usually right.
     */
    function draw() {
        try {
            // Handlers FIRST, then the tab. That is the order ONLYOFFICE's own helper uses in
            // plugins-ui.js — it attaches every button's onclick and only then calls
            // AddToolbarMenuItem. Attaching afterwards left every button doing nothing at all:
            // the tab drew, the clicks went nowhere, and no request ever reached the service.
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
            // Nothing here can show a message — there is no window. The tab simply does not
            // appear, and the panel still works.
            window.console && window.console.log("Nexus PLM: the tab could not be drawn — " + e.message);
        }
    }

    /** Ask the service who is signed in and what the open document is, then redraw the tab. */
    function refresh() {
        client.me().then(function (me) {
            context.signedIn = !!(me && me.success && me.username);
            context.user = context.signedIn ? me.username : null;

            if (!context.signedIn) { context.state = null; draw(); return; }

            // Which file the editor has open is not known yet, so nothing that needs the document
            // is offered. That is honest rather than optimistic: the commands would be refused.
            client.state({}).then(function (state) {
                context.state = state;
                draw();
            });
        });
    }

    function run(commandId) {
        var command = commands.byId(commandId);
        if (!command) { return; }

        // Never act on a command the rules say is unavailable. The tab greys it, but a stale tab
        // is one repaint away, and the service's refusal is indistinguishable from a bug.
        if (!commands.isEnabled(command, context)) {
            say(commands.disabledBecause(command, context));
            return;
        }

        if (command.id === "navigator") { showPanel(); return; }
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

    /** Say something through the tray's own toast, so it looks like it does in every host. */
    function say(message, severity) {
        if (!message) { return; }
        try { client.notify(message, severity || "info"); }
        catch (e) { window.console && window.console.log(message); }
    }

    /** Connection Status has no dialog of its own: it reports what we can see from here. */
    function reportConnection() {
        client.health().then(function (answer) {
            if (answer && answer.unreachable) {
                say(answer.error, "error");
            } else if (context.signedIn) {
                say("Connected to Nexus PLM, signed in as " + context.user + ".", "success");
            } else {
                say("Connected to Nexus PLM. Not signed in.", "info");
            }
        });
    }

    /** Open the panel — this variation has no window of its own to show. */
    function showPanel() {
        try { window.Asc.plugin.executeMethod("ShowPlugin", [window.Asc.plugin.guid]); }
        catch (e) { window.console && window.console.log("Nexus PLM: could not open the panel."); }
    }

    function openHelp() {
        try { window.Asc.plugin.executeMethod("OpenLink", ["https://nexusplm.help/"]); }
        catch (e) { window.console && window.console.log("Nexus PLM: could not open help."); }
    }

})(window);
