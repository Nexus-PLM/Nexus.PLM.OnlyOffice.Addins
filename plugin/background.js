/*
 * The Nexus PLM tab.
 *
 * This is the plugin's resident half: the frame ONLYOFFICE keeps alive for as long as the
 * document is open, because config.json types its variation "background". It puts the tab in
 * the ribbon, keeps every button's enabled state in step with the signed-in user and the open
 * document, and hands what the user presses to the runner — the whole of what the Word ribbon
 * does, through the same service.
 *
 * ── Why "background", and not the invisible variation this used to be ─────────────────────
 * The SDK maps a variation with no `type` and `isVisual: false` to PluginType.Invisible: a
 * one-shot action whose frame is torn down as soon as init returns. The tab it registered
 * survived that, which is why buttons drew and then did nothing — there was no frame left for a
 * click to reach. A `type: "background"` variation is what the shipped AI plugin uses to keep
 * its tab alive, and the editor starts it with every document unless the user stops it. (Read
 * from CPluginVariation in sdk-all-min.js, Desktop Editors 9.4.0.)
 */

(function (window) {
    "use strict";

    var toolbar = window.NexusPlmToolbar;
    var clientModule = window.NexusPlmClient;
    var editor = window.NexusPlmEditor;

    /** How often to re-read the document's state when nothing else has, as Word's ribbon does. */
    var POLL_MS = 15000;

    var runner = null;

    /** True once the tab has been drawn; afterwards it is updated, never re-added. */
    var tabDrawn = false;

    var beat = null;
    var poll = null;

    window.Asc = window.Asc || {};
    window.Asc.plugin = window.Asc.plugin || {};

    window.Asc.plugin.init = function () {
        runner = window.NexusPlmRunner.create({
            editorType: (window.Asc.plugin.info && window.Asc.plugin.info.editorType) || null,
            onState: drawTab
        });

        // Register with the tray first, so Nexus lists this host as connected the way it lists
        // Word and FreeCAD, and keep telling it we are here.
        runner.client.connect();
        beat = window.setInterval(function () { runner.client.heartbeat(); }, clientModule.HEARTBEAT_MS);

        attachHandlers();
        drawTab(runner.context);
        runner.refresh();
        poll = window.setInterval(function () { runner.refresh(); }, POLL_MS);
    };

    window.addEventListener("unload", function () {
        if (beat) { window.clearInterval(beat); beat = null; }
        if (poll) { window.clearInterval(poll); poll = null; }
        if (runner) { runner.client.disconnect(); }
    });

    /** The document changed under us; what the commands may do may have changed too. */
    window.Asc.plugin.onDocumentContentReady = function () { if (runner) { runner.refresh(); } };

    /** A window of ours closing arrives here with id -1. */
    window.Asc.plugin.button = function (id, windowId) {
        if (id === -1) { editor.windowClosed(windowId); }
    };

    window.Asc.plugin.onExternalMouseUp = function () {};

    /**
     * Handlers FIRST, then the tab — the order ONLYOFFICE's own plugins-ui.js uses. Every button
     * and every entry in a button's menu has an id of its own, and a click arrives under it.
     */
    function attachHandlers() {
        toolbar.clickIds().forEach(function (id) {
            window.Asc.plugin.attachToolbarMenuClickEvent(id, function () {
                runner.run(toolbar.commandFor(id));
            });
        });
    }

    /**
     * Add the tab once; afterwards send the same payload as an update, which the editor applies
     * to the buttons it already drew (text, tooltip, enabled state, menus).
     */
    function drawTab(context) {
        try {
            window.Asc.plugin.executeMethod(
                tabDrawn ? "UpdateToolbarMenuItem" : "AddToolbarMenuItem",
                [toolbar.tab(window.Asc.plugin.guid, context)]);
            tabDrawn = true;
        } catch (e) {
            window.console && window.console.log("Nexus PLM: the tab could not be drawn - " + e.message);
        }
    }

})(window);
