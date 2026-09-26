/*
 * What the Navigator says when it cannot show the vault, and what it calls itself.
 *
 * Small on purpose: the pane's own content — the tree, the rows, the actions — is decided in
 * `navigator.js`. What is left here is the handful of sentences that are about the connection
 * rather than about the vault, and the pane's title. They live apart because each one is the
 * answer to a different failure, and reporting all of them as "sign in" was wrong: only one of
 * them is fixed by signing in.
 */

(function (root, factory) {
    if (typeof module === "object" && module.exports) {
        module.exports = factory(require("./editors.js"));
    } else {
        root.NexusPlmPanel = factory(root.NexusPlmEditors);
    }
}(typeof self !== "undefined" ? self : this, function (editors) {
    "use strict";

    /**
     * Shown when the plugin has no per-install secret, so the service refuses it as a caller it
     * cannot show to be local. Named as a setup problem, because that is what it is: somebody
     * copied the plugin into place without running the install step. A bare 403 would send them
     * looking at permissions in PLM.
     */
    var NOT_INSTALLED = "This plugin was not installed properly: run the install step so it can "
        + "prove it is running on your machine.";

    /** The service is not answering at all — the one failure the user has to fix themselves. */
    var NOT_RUNNING = "Nexus PLM is not running. Start the Nexus PLM Addins tray application.";

    /** The service answered, and would not deal with us. Different from being signed out. */
    var REFUSED = "Nexus PLM refused this plugin. Re-run the install step to give it the current "
        + "secret for this machine.";

    /** No session. The pane says where to fix it, because Sign In is on the tab, not in here. */
    var NOT_SIGNED_IN = "You are not signed in to Nexus PLM. Use Sign In on the Nexus PLM tab.";

    /**
     * The one line to show instead of the vault, or null when there is nothing wrong.
     *
     * Order matters: with no secret every call is refused, so any other message would be a guess
     * about a service that never answered.
     */
    function trouble(options) {
        var o = options || {};
        if (o.hasSecret === false) { return NOT_INSTALLED; }
        if (o.unreachable) { return NOT_RUNNING; }
        if (o.refused) { return REFUSED; }
        if (o.signedIn === false) { return NOT_SIGNED_IN; }
        return o.error || null;
    }

    /**
     * The pane's title, naming the editor so a user can tell the three apart when two are open.
     * Falls back to the plain name when ONLYOFFICE reports an editor we do not know.
     */
    function title(editorType) {
        var e = editors.describe(editorType);
        return e ? ("Nexus PLM — " + e.name) : "Nexus PLM";
    }

    return {
        NOT_INSTALLED: NOT_INSTALLED,
        NOT_RUNNING: NOT_RUNNING,
        REFUSED: REFUSED,
        NOT_SIGNED_IN: NOT_SIGNED_IN,
        trouble: trouble,
        title: title
    };
}));
