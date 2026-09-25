/*
 * What the panel shows, decided without ONLYOFFICE anywhere near it.
 *
 * Every decision about rows, buttons, and what to say when PLM has never seen the document lives
 * here and is tested directly. The LibreOffice add-in keeps the same split (`nexusplm/panel.py`)
 * and it is the reason its panel could be changed repeatedly without an editor to hand.
 */

(function (root, factory) {
    if (typeof module === "object" && module.exports) {
        module.exports = factory(require("./editors.js"));
    } else {
        root.NexusPlmPanel = factory(root.NexusPlmEditors);
    }
}(typeof self !== "undefined" ? self : this, function (editors) {
    "use strict";

    /** Shown in place of a value the server did not send. An empty cell reads as a bug. */
    var ABSENT = "—";

    /** The ordinary case for a file somebody just created, so it is phrased as a fact. */
    var NOT_IN_PLM = "This document is not in PLM.";
    var NO_DOCUMENT = "No document is open.";
    var NOT_SIGNED_IN = "Sign in to Nexus PLM to see this document.";

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
     * Each button: label, the command it runs, and when it may be pressed.
     *
     * The rule matches the LibreOffice add-in's and the Office ribbon's, because a user who finds
     * a command enabled in one host and greyed in another has found a bug, whichever is right.
     */
    var BUTTONS = [
        ["Check Out", "check_out", "checked_in"],
        ["Check In", "check_in", "mine"],
        ["Save to PLM", "save_to_plm", "mine"],
        ["Edit Values", "edit_values", "in_plm"],
        ["Refresh", "refresh_values", "in_plm"]
    ];

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
        if (o.signedIn === false) { return NOT_SIGNED_IN; }
        if (o.hasDocument === false) { return NO_DOCUMENT; }
        if (!isInPlm(state)) { return NOT_IN_PLM; }
        return state.part_number || ABSENT;
    }

    /**
     * Which buttons may be pressed for this state, as an array of command names.
     *
     * `user` is who is signed in: a document checked out to somebody else is not yours to check
     * in, and the panel must not offer it — the service would refuse, and an offer that is always
     * refused is worse than no offer.
     */
    function enabledButtons(state, user) {
        if (!isInPlm(state)) { return []; }
        var status = state.status;
        var mine = status === "checked_out" && !!user && state.checked_out_by === user;

        return BUTTONS.filter(function (b) {
            var rule = b[2];
            if (rule === "in_plm") { return true; }
            if (rule === "checked_in") { return status === "checked_in"; }
            if (rule === "mine") { return mine; }
            return false;
        }).map(function (b) { return b[1]; });
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
        ROWS: ROWS,
        BUTTONS: BUTTONS,
        isInPlm: isInPlm,
        rowsFor: rowsFor,
        headline: headline,
        enabledButtons: enabledButtons,
        title: title
    };
}));
