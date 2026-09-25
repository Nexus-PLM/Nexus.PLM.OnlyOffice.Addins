/*
 * The Nexus PLM tab: its groups, its commands, and what each of them needs.
 *
 * ── Modelled on the Word ribbon, deliberately ────────────────────────────────────────────────
 * The tab name, the group names, their order and the labels are `PLMRibbon.xml` in
 * `Nexus.PLM.Office.WordAddin`, read from the file rather than remembered. A user who has used
 * Nexus in Word should not have to learn a second arrangement to use it here, and a command that
 * is called "Save As Existing Item" in one host and something else in another is a support call.
 *
 * Word's set is larger than LibreOffice's: it has the Navigator, and it has Markup and Apply
 * Markups, which LibreOffice never got.
 *
 * One table, because the toolbar, the panel and the dispatcher must never disagree about what a
 * command is called, where it lives or when it may be pressed. A user who finds a command enabled
 * in one place and greyed in another has found a bug, whichever one is right.
 *
 * The endpoints and the "needs a document" column are not guesses either: they are taken from the
 * LibreOffice client, and from which of its handlers call `_require_item` / `_require_path`.
 */

(function (root, factory) {
    if (typeof module === "object" && module.exports) { module.exports = factory(); }
    else { root.NexusPlmCommands = factory(); }
}(typeof self !== "undefined" ? self : this, function () {
    "use strict";

    /** The tab's name, as Word calls it. */
    var TAB = "Nexus PLM";

    /**
     * When a command may be pressed.
     *
     * ALWAYS     — even signed out, even with nothing open. Only things that report or explain.
     * SIGNED_OUT — only when there is no session.
     * SIGNED_IN  — needs a session, but nothing about the document.
     * DOCUMENT   — needs to know which PLM item the open file is.
     * CHECKED_IN — that, and the item not checked out.
     * MINE       — that, and checked out to the signed-in user.
     */
    var ALWAYS = "always";
    var SIGNED_OUT = "signed_out";
    var SIGNED_IN = "signed_in";
    var DOCUMENT = "document";
    var CHECKED_IN = "checked_in";
    var MINE = "mine";

    /** Groups, in Word's order. */
    var GROUPS = ["Account", "Data Management", "Navigation and View", "Tasks",
                  "Workflow", "Attribute Exchange", "Settings"];

    /**
     * Every command. `host: false` marks one whose service endpoint exists but whose half in
     * *this* host does not yet — it is shown, and says so, rather than being quietly missing.
     */
    var COMMANDS = [
        // ── Account ──────────────────────────────────────────────────────────
        { id: "sign_in",          label: "Sign In",               group: "Account",             endpoint: "/api/auth/login",       when: SIGNED_OUT },
        { id: "sign_out",         label: "Sign Out",              group: "Account",             endpoint: "/api/auth/logout",      when: SIGNED_IN },

        // ── Data Management ──────────────────────────────────────────────────
        { id: "new",              label: "New",                   group: "Data Management",     endpoint: "/plm/new",              when: SIGNED_IN },
        { id: "open",             label: "Open",                  group: "Data Management",     endpoint: "/plm/open",             when: SIGNED_IN },
        { id: "save",             label: "Save",                  group: "Data Management",     endpoint: "/plm/save",             when: MINE },
        { id: "save_as_new",      label: "Save As",               group: "Data Management",     endpoint: "/plm/save-as-new",      when: SIGNED_IN },
        { id: "save_as_existing", label: "Save As Existing",      group: "Data Management",     endpoint: "/plm/save-as-existing", when: SIGNED_IN },

        // ── Navigation and View ──────────────────────────────────────────────
        //: The Navigator is this plugin's own panel, not a service call — it is already built.
        { id: "navigator",        label: "Navigator",             group: "Navigation and View", endpoint: null,                    when: ALWAYS, toggle: true },
        { id: "search",           label: "Search",                group: "Navigation and View", endpoint: "/plm/search",           when: SIGNED_IN },
        { id: "properties",       label: "Properties",            group: "Navigation and View", endpoint: "/plm/properties",       when: DOCUMENT },

        // ── Tasks ────────────────────────────────────────────────────────────
        //: Word shows Check Out and Check In as one split button whose label follows the state.
        //: Here they are two commands and the enable rules do the same job, which is the same
        //: behaviour without a control that lies about which half you pressed.
        { id: "check_out",        label: "Check Out",             group: "Tasks",               endpoint: "/plm/checkout",         when: CHECKED_IN },
        { id: "check_in",         label: "Check In",              group: "Tasks",               endpoint: "/plm/checkin",          when: MINE },
        { id: "release",          label: "Release",               group: "Tasks",               endpoint: "/plm/release",          when: DOCUMENT },
        { id: "revise",           label: "Revise",                group: "Tasks",               endpoint: "/plm/revise",           when: DOCUMENT },
        { id: "markup",           label: "Markup",                group: "Tasks",               endpoint: "/plm/markup",           when: DOCUMENT, host: false },
        { id: "apply_markups",    label: "Apply markups from PLM", group: "Tasks",              endpoint: "/plm/markup",           when: DOCUMENT, host: false },
        { id: "change_owner",     label: "Change Ownership",      group: "Tasks",               endpoint: "/plm/set-owner",        when: DOCUMENT },

        // ── Workflow ─────────────────────────────────────────────────────────
        { id: "worklist",         label: "My Worklist",           group: "Workflow",            endpoint: "/plm/worklist",         when: SIGNED_IN },
        { id: "new_workflow",     label: "New Workflow",          group: "Workflow",            endpoint: "/plm/workflow",         when: DOCUMENT },

        // ── Attribute Exchange ───────────────────────────────────────────────
        { id: "edit_values",      label: "Edit Values",           group: "Attribute Exchange",  endpoint: "/plm/edit-values",      when: DOCUMENT },
        { id: "refresh_values",   label: "Refresh Values",        group: "Attribute Exchange",  endpoint: "/plm/refresh-values",   when: DOCUMENT },
        { id: "reload_document",  label: "Reload Document",       group: "Attribute Exchange",  endpoint: "/plm/reload-document",  when: DOCUMENT },

        // ── Settings ─────────────────────────────────────────────────────────
        { id: "settings",         label: "Current Settings",      group: "Settings",            endpoint: "/plm/settings",         when: SIGNED_IN },
        { id: "help",             label: "Help",                  group: "Settings",            endpoint: null,                    when: ALWAYS },
        { id: "about",            label: "About Nexus PLM",       group: "Settings",            endpoint: "/plm/about",            when: ALWAYS },
        { id: "connection",       label: "Connection Status",     group: "Settings",            endpoint: "/api/health",           when: ALWAYS }
    ];

    /**
     * Whether a command cannot work without knowing which PLM item the open file is.
     *
     * Derived from `when` rather than stored beside it, so the two can never disagree — which a
     * second column eventually always does.
     */
    function needsDocument(command) {
        return command.when === DOCUMENT || command.when === CHECKED_IN || command.when === MINE;
    }

    /** Whether this host can actually carry the command out yet. */
    function hostSupports(command) {
        return command.host !== false;
    }

    function byId(id) {
        for (var i = 0; i < COMMANDS.length; i++) {
            if (COMMANDS[i].id === id) { return COMMANDS[i]; }
        }
        return null;
    }

    function inGroup(group) {
        return COMMANDS.filter(function (c) { return c.group === group; });
    }

    /**
     * Whether this command may be pressed right now.
     *
     * `context` is `{ signedIn, user, state }` — the same three things the panel draws from, so
     * the tab and the panel cannot reach different answers.
     */
    function isEnabled(command, context) {
        if (!hostSupports(command)) { return false; }

        var c = context || {};
        var state = c.state;
        var inPlm = !!(state && state.status && state.status !== "unknown");
        var mine = inPlm && state.status === "checked_out" && !!c.user &&
                   state.checked_out_by === c.user;

        switch (command.when) {
            case ALWAYS:     return true;
            case SIGNED_OUT: return !c.signedIn;
            case SIGNED_IN:  return !!c.signedIn;
            case DOCUMENT:   return !!c.signedIn && inPlm;
            case CHECKED_IN: return !!c.signedIn && inPlm && state.status === "checked_in";
            case MINE:       return !!c.signedIn && mine;
            default:         return false;
        }
    }

    /**
     * Why a command is greyed, for its tooltip — so a disabled button is never a dead end.
     * Null when it is not greyed.
     */
    function disabledBecause(command, context) {
        if (isEnabled(command, context)) { return null; }
        var c = context || {};

        if (!hostSupports(command)) {
            return "Nexus PLM does not read comments from ONLYOFFICE yet.";
        }
        if (command.when === SIGNED_OUT) { return "You are already signed in."; }
        if (!c.signedIn) { return "Sign in to Nexus PLM first."; }

        var state = c.state;
        var inPlm = !!(state && state.status && state.status !== "unknown");
        if (needsDocument(command) && !inPlm) {
            return "Nexus PLM does not know which item this document is.";
        }
        if (command.when === CHECKED_IN) { return "This document is already checked out."; }
        if (command.when === MINE) { return "This document is not checked out to you."; }
        return "Not available for this document.";
    }

    return {
        TAB: TAB,
        ALWAYS: ALWAYS, SIGNED_OUT: SIGNED_OUT, SIGNED_IN: SIGNED_IN,
        DOCUMENT: DOCUMENT, CHECKED_IN: CHECKED_IN, MINE: MINE,
        COMMANDS: COMMANDS,
        GROUPS: GROUPS,
        needsDocument: needsDocument,
        hostSupports: hostSupports,
        byId: byId,
        inGroup: inGroup,
        isEnabled: isEnabled,
        disabledBecause: disabledBecause
    };
}));
