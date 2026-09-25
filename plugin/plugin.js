/*
 * The plugin's entry point: what ONLYOFFICE calls, and what it does about it.
 *
 * ONE plugin serves all three editors. `EditorsSupport` in config.json names word, cell and
 * slide, and nothing in here is editor-specific — the editor is asked what it is at init, and the
 * few places that differ (what "the document" is called, which fields a spreadsheet can drive)
 * branch on that one answer. The LibreOffice add-in is built the same way for its five
 * applications, and the thing that keeps it honest is that there is no second copy to drift.
 *
 * Everything that can be decided without an editor is decided in the plain modules under `lib/`,
 * which the tests exercise directly. `ui.js` puts those decisions into the DOM. What is left in
 * this file is only the glue that cannot be tested without ONLYOFFICE: the Asc.plugin callbacks.
 */

(function (window) {
    "use strict";

    /** Which editor we are in, as ONLYOFFICE names it: "word" | "cell" | "slide". */
    var editor = null;

    window.Asc = window.Asc || {};
    window.Asc.plugin = window.Asc.plugin || {};

    /**
     * Called once when the editor loads the plugin.
     *
     * `Asc.plugin.info.editorType` is how the plugin learns which of the three it is in. It is
     * read here and nowhere else, so a new editor means one entry in config.json and one case in
     * `lib/editors.js`, not a new copy of the plugin.
     */
    window.Asc.plugin.init = function () {
        editor = (window.Asc.plugin.info && window.Asc.plugin.info.editorType) || null;
        window.NexusPlmUi.start(editor);
    };

    /**
     * Called when the user presses a button in a modal variation, and when the plugin is closed.
     * `id === -1` is the close box; every other id indexes `buttons` in config.json.
     */
    window.Asc.plugin.button = function (id) {
        if (id === -1) {
            window.Asc.plugin.executeCommand("close", "");
        }
    };

    /**
     * Called when the document's contents change under us.
     *
     * The panel shows PLM's view of the item, not the document's, so a keystroke must not cause a
     * round trip to the service. This only marks the panel stale.
     */
    window.Asc.plugin.onDocumentContentReady = function () {
        window.NexusPlmUi.documentReady();
    };

    /** ONLYOFFICE calls this when the plugin is being torn down. */
    window.Asc.plugin.onExternalMouseUp = function () {};

})(window);
