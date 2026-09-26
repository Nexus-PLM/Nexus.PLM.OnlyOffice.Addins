/*
 * What the panel shows, decided without ONLYOFFICE anywhere near it.
 *
 * Every decision about rows, buttons, and what to say when PLM has never seen the document lives
 * here and is tested directly. The LibreOffice add-in keeps the same split (`nexusplm/panel.py`)
 * and it is the reason its panel could be changed repeatedly without an editor to hand.
 */

(function (root, factory) {
    if (typeof module === "object" && module.exports) {
        module.exports = factory(require("./editors.js"), require("./commands.js"));
    } else {
        root.NexusPlmPanel = factory(root.NexusPlmEditors, root.NexusPlmCommands);
    }
}(typeof self !== "undefined" ? self : this, function (editors, commands) {
    "use strict";

    /** Shown in place of a value the server did not send. An empty cell reads as a bug. */
    var ABSENT = "—";

    /** The ordinary case for a file somebody just created, so it is phrased as a fact. */
    var NOT_IN_PLM = "This document is not in PLM.";
    var NO_DOCUMENT = "No document is open.";
    var NOT_SIGNED_IN = "Sign in to Nexus PLM to see this document.";

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

    /** (label, the key on a /plm/state answer). */
    var ROWS = [
        ["Part number", "part_number"],
        ["Revision", "revision"],
        ["Status", "status"],
        ["Checked out", "checked_out_by"],
        ["Type", "type_name"],
        ["Description", "description"]
    ];

    /**
     * The handful of commands the panel puts within reach, as [label, id] pairs.
     *
     * Only the ids are chosen here. The label and the rule for when each may be pressed come from
     * the one command table, so the panel and the Nexus PLM tab cannot disagree — a user who finds
     * a command enabled in one and greyed in the other has found a bug, whichever is right.
     */
    var PANEL_COMMANDS = ["check_out", "check_in", "save", "edit_values", "refresh_values"];

    var BUTTONS = PANEL_COMMANDS.map(function (id) {
        var command = commands.byId(id);
        return [command.label, command.id, command.when];
    });

    function statusWord(status) {
        if (!status) { return ABSENT; }
        return String(status).replace(/_/g, " ").replace(/^./, function (c) { return c.toUpperCase(); });
    }

    /**
     * Whether the answer describes a document PLM actually knows.
     *
     * `status` is "unknown" both for a file PLM has never seen and for one it could not resolve,
     * so this is deliberately about what can be shown, not about why.
     */
    function isInPlm(state) {
        return !!(state && state.status && state.status !== "unknown");
    }

    /** The (label, value) pairs to display for a /plm/state answer. */
    function rowsFor(state) {
        if (!isInPlm(state)) {
            return ROWS.map(function (r) { return [r[0], ABSENT]; });
        }
        return ROWS.map(function (r) {
            var value = state[r[1]];
            if (r[1] === "status") { value = statusWord(value); }
            return [r[0], (value === undefined || value === null || value === "") ? ABSENT : String(value)];
        });
    }

    /** The one line at the top of the panel, above the rows. */
    function headline(state, options) {
        var o = options || {};
        // Before anything else: with no secret every call is refused, so any other message would
        // be a guess about a service that never answered.
        if (o.hasSecret === false) { return NOT_INSTALLED; }
        if (o.unreachable) { return NOT_RUNNING; }
        if (o.refused) { return REFUSED; }
        if (o.signedIn === false) { return NOT_SIGNED_IN; }
        if (o.hasDocument === false) { return NO_DOCUMENT; }
        if (!isInPlm(state)) { return NOT_IN_PLM; }
        return state.part_number || ABSENT;
    }

    /**
     * Which of the panel's buttons may be pressed, as an array of command ids.
     *
     * Delegated to the command table rather than decided again here. The panel is only ever
     * reached with a session, so `signedIn` is implied by having a user.
     */
    function enabledButtons(state, user) {
        var context = { signedIn: !!user, user: user, state: state };
        return PANEL_COMMANDS.filter(function (id) {
            return commands.isEnabled(commands.byId(id), context);
        });
    }

    /**
     * The heading over the panel, naming the editor so a user can tell the three apart when two
     * are open. Falls back to the plain name when ONLYOFFICE reports an editor we do not know.
     */
    function title(editorType) {
        var e = editors.describe(editorType);
        return e ? ("Nexus PLM — " + e.name) : "Nexus PLM";
    }

    return {
        ABSENT: ABSENT,
        NOT_IN_PLM: NOT_IN_PLM,
        NO_DOCUMENT: NO_DOCUMENT,
        NOT_SIGNED_IN: NOT_SIGNED_IN,
        NOT_INSTALLED: NOT_INSTALLED,
        NOT_RUNNING: NOT_RUNNING,
        REFUSED: REFUSED,
        ROWS: ROWS,
        BUTTONS: BUTTONS,
        PANEL_COMMANDS: PANEL_COMMANDS,
        isInPlm: isInPlm,
        rowsFor: rowsFor,
        headline: headline,
        enabledButtons: enabledButtons,
        title: title
    };
}));
