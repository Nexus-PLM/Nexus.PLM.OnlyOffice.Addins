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

    window.Asc = window.Asc || {};
    window.Asc.plugin = window.Asc.plugin || {};

    window.Asc.plugin.init = function () {
        draw();
        refresh();
    };

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
            window.Asc.plugin.executeMethod(
                "AddToolbarMenuItem", [toolbar.tab(window.Asc.plugin.guid, context)]);

            if (!handlersAttached) {
                commands.COMMANDS.forEach(function (command) {
                    window.Asc.plugin.attachToolbarMenuClickEvent(
                        toolbar.buttonId(command.id),
                        function () { run(command.id); });
                });
                handlersAttached = true;
            }
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

        var body = { hwnd: 0 };
        if (commands.needsDocument(command) && context.state) {
            body.item_id = context.state.item_id;
        }

        client.command(command.endpoint, body).then(function (answer) {
            if (answer && answer.success === false && answer.error) { say(answer.error); }
            refresh();
        });
    }

    /** Say something through the service's own toast, so it looks like it does in every host. */
    function say(message) {
        if (!message) { return; }
        try { client.notify(message, "warning"); }
        catch (e) { window.console && window.console.log(message); }
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
