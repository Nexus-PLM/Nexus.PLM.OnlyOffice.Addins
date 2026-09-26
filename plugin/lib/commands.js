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
    if (typeof module === "object" && module.exports) {
        module.exports = factory(require("./client.js"));
    } else {
        root.NexusPlmCommands = factory(root.NexusPlmClient);
    }
}(typeof self !== "undefined" ? self : this, function (client) {
    "use strict";

    /** The tab's name, as Word calls it. */
    var TAB = "Nexus PLM";

    /**
     * When a command may be pressed.
     *
     * ALWAYS     — even signed out, even with nothing open. Only things that report or explain.
     * SIGNED_OUT — only when there is no session.
     * SIGNED_IN  — needs a session to be offered at all. Only Sign Out.
     * SESSION    — offered always; the press signs you in first, as Word's handlers do.
     * UNREGISTERED — needs a session, and the document must not already be a PLM item.
     * DOCUMENT   — needs to know which PLM item the open file is.
     * CHECKED_IN — that, and the item not checked out.
     * MINE       — that, and checked out to the signed-in user.
     * NOT_MINE   — that, and NOT checked out to the signed-in user: what Word requires before
     *              replacing the document with PLM's copy, because a copy checked out to you is
     *              the newer one.
     */
    var ALWAYS = "always";
    var SIGNED_OUT = "signed_out";
    var SIGNED_IN = "signed_in";
    /**
     * Pressable whether or not there is a session: the press signs you in and then runs, which
     * is what Word does (every handler calls EnsureSignedInAsync first). Nothing about the
     * document is needed either - these are how you GET a document.
     */
    var SESSION = "session";
    /** Needs a session, and the open document must NOT already be a PLM item. */
    var UNREGISTERED = "unregistered";
    var DOCUMENT = "document";
    var CHECKED_IN = "checked_in";
    var MINE = "mine";
    /** Needs the document, and it must NOT be checked out to the signed-in user. */
    var NOT_MINE = "not_mine";

    /** Groups, in Word's order. */
    var GROUPS = ["Account", "Data Management", "Navigation and View", "Tasks",
                  "Workflow", "Attribute Exchange", "Settings"];

    /**
     * Every command.
     *
     * `parent` puts a command in another's menu, the way Word's split buttons carry Release,
     * Apply markups, Reload Document, About and Connection Status. `saves` marks one that
     * uploads the file on disk. `method` is POST unless said otherwise.
     */
    var COMMANDS = [
        // ── Account ──────────────────────────────────────────────────────────
        { id: "sign_in",          label: "Sign In",               group: "Account",             endpoint: "/api/auth/login",       when: SIGNED_OUT },
        { id: "sign_out",         label: "Sign Out",              group: "Account",             endpoint: "/api/auth/logout",      when: SIGNED_IN },

        // ── Data Management ──────────────────────────────────────────────────
        //: `saves` marks a command that uploads the file on disk, so the host must have the
        //: document written out first. Word saves it itself; see host.js for what this one does.
        { id: "new",              label: "New",                   group: "Data Management",     endpoint: "/plm/new",              when: SESSION },
        { id: "open",             label: "Open",                  group: "Data Management",     endpoint: "/plm/open",             when: SESSION },
        { id: "save",             label: "Save",                  group: "Data Management",     endpoint: "/plm/save",             when: MINE, saves: true },
        //: Word refuses this outright once the document IS an item, in the ribbon and again in
        //: the handler ("already registered in PLM - use Save to update it"). Offering it would
        //: create a SECOND item for one file.
        { id: "save_as_new",      label: "Save As",               group: "Data Management",     endpoint: "/plm/save-as-new",      when: UNREGISTERED, saves: true },
        { id: "save_as_existing", label: "Save As Existing",      group: "Data Management",     endpoint: "/plm/save-as-existing", when: SIGNED_IN, saves: true },

        // ── Navigation and View ──────────────────────────────────────────────
        //: The Navigator is this plugin's own panel, not a service call — it is already built.
        { id: "navigator",        label: "Navigator",             group: "Navigation and View", endpoint: null,                    when: ALWAYS, toggle: true },
        { id: "search",           label: "Search",                group: "Navigation and View", endpoint: "/plm/search",           when: SESSION },
        { id: "properties",       label: "Properties",            group: "Navigation and View", endpoint: "/plm/properties",       when: DOCUMENT },

        // ── Tasks ────────────────────────────────────────────────────────────
        //: Word shows Check Out and Check In as one split button whose label follows the state,
        //: with Release and Revise in its menu. Here every one is its own button: ONLYOFFICE greys
        //: a split button's menu with the button, and Release must stay reachable while Check In
        //: is not (it needs the document checked IN). The enable rules do the same job as Word's
        //: label switch without a control that lies about which half you pressed.
        { id: "check_out",        label: "Check Out",             group: "Tasks",               endpoint: "/plm/checkout",         when: CHECKED_IN },
        { id: "check_in",         label: "Check In",              group: "Tasks",               endpoint: "/plm/checkin",          when: MINE, saves: true },
        { id: "release",          label: "Release",               group: "Tasks",               endpoint: "/plm/release",          when: CHECKED_IN },
        { id: "revise",           label: "Revise",                group: "Tasks",               endpoint: "/plm/revise",           when: DOCUMENT },
        //: Markup sends this document's comments to PLM as review markups; its menu brings every
        //: reviewer's markups back in as comments — the same pair as Word's split button.
        { id: "markup",           label: "Markup",                group: "Tasks",               endpoint: "/plm/markup",           when: DOCUMENT },
        { id: "apply_markups",    label: "Apply markups from PLM", group: "Tasks",              endpoint: "/plm/markup",           when: DOCUMENT, parent: "markup", method: "GET" },
        { id: "change_owner",     label: "Change Ownership",      group: "Tasks",               endpoint: "/plm/set-owner",        when: DOCUMENT, saves: true },

        // ── Workflow ─────────────────────────────────────────────────────────
        { id: "worklist",         label: "My Worklist",           group: "Workflow",            endpoint: "/plm/worklist",         when: SESSION },
        //: Word refuses a checked-out document ("check in first"), so it is greyed here for the
        //: same reason rather than offered and refused.
        { id: "new_workflow",     label: "New Workflow",          group: "Workflow",            endpoint: "/plm/workflow",         when: CHECKED_IN },

        // ── Attribute Exchange ───────────────────────────────────────────────
        //: Word: "Check out the document before editing attribute values."
        { id: "edit_values",      label: "Edit Values",           group: "Attribute Exchange",  endpoint: "/plm/edit-values",      when: MINE },
        { id: "refresh_values",   label: "Refresh Values",        group: "Attribute Exchange",  endpoint: "/plm/refresh-values",   when: DOCUMENT },
        //: Word: "You have this document checked out - your copy is the newer one."
        { id: "reload_document",  label: "Reload Document",       group: "Attribute Exchange",  endpoint: "/plm/reload-document",  when: NOT_MINE, parent: "refresh_values" },

        // ── Settings ─────────────────────────────────────────────────────────
        { id: "settings",         label: "Current Settings",      group: "Settings",            endpoint: "/plm/settings",         when: SESSION },
        { id: "help",             label: "Help",                  group: "Settings",            endpoint: null,                    when: ALWAYS },
        { id: "about",            label: "About Nexus PLM",       group: "Settings",            endpoint: "/plm/about",            when: ALWAYS, parent: "help" },
        { id: "connection",       label: "Connection Status",     group: "Settings",            endpoint: "/api/health",           when: ALWAYS, parent: "help" }
    ];

    /**
     * Whether a command cannot work without knowing which PLM item the open file is.
     *
     * Derived from `when` rather than stored beside it, so the two can never disagree — which a
     * second column eventually always does.
     */
    function needsDocument(command) {
        return command.when === DOCUMENT || command.when === CHECKED_IN || command.when === MINE ||
               command.when === NOT_MINE;
    }

    /**
     * Whether pressing this command without a session should climb the sign-in ladder rather
     * than refuse. Word does this for every command: each handler calls EnsureSignedInAsync
     * first, so a signed-out user presses New once, signs in, and New runs.
     */
    function needsSession(command) {
        return !!command && command.when !== ALWAYS && command.when !== SIGNED_OUT;
    }

    /** Every command, as the tab draws them, needs a session for its rule to mean anything. */
    function sessionOptional(command) {
        return !!command && command.when === SESSION;
    }

    /** The commands drawn as buttons: everything that is not in another command's menu. */
    function topLevel() {
        return COMMANDS.filter(function (c) { return !c.parent; });
    }

    /** The commands in this one's menu, in table order. */
    function childrenOf(commandId) {
        return COMMANDS.filter(function (c) { return c.parent === commandId; });
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
        var c = context || {};
        var state = c.state;
        var inPlm = !!(state && state.status && state.status !== "unknown");
        var mine = inPlm && state.status === "checked_out" && !!c.user &&
                   state.checked_out_by === c.user;

        switch (command.when) {
            case ALWAYS:     return true;
            case SIGNED_OUT: return !c.signedIn;
            case SIGNED_IN:  return !!c.signedIn;
            //: Offered signed out too - pressing it climbs the sign-in ladder, then runs.
            case SESSION:    return true;
            case UNREGISTERED: return !!c.signedIn && !inPlm;
            case DOCUMENT:   return !!c.signedIn && inPlm;
            case CHECKED_IN: return !!c.signedIn && inPlm && state.status === "checked_in";
            case MINE:       return !!c.signedIn && mine;
            case NOT_MINE:   return !!c.signedIn && inPlm && !mine;
            default:         return false;
        }
    }

    /**
     * What to POST for this command.
     *
     * Not one shape for all of them: the browsers need to know what this host can open and what
     * to root themselves at, and the item commands need the item. Taken from the WORD add-in,
     * which is the most complete one and the one these endpoints were shaped around - a command
     * that posts the wrong body is answered, refused, and looks exactly like a command that did
     * nothing at all.
     */
    function bodyFor(command, context) {
        var c = context || {};
        var itemId = (c.state && c.state.item_id) || null;
        // Where the file is comes from the host's own reading of the document (`doc`), not from
        // PLM's answer about the item: a /plm/state answer carries no path.
        var filePath = (c.doc && c.doc.path) || (c.state && c.state.file_path) || null;
        var extensions = client.FILE_EXTENSIONS.join(";");

        switch (command.id) {
            case "sign_in":          return { hwnd: 0 };
            case "sign_out":         return {};

            case "new":              return { root_base_type: client.ROOT_BASE_TYPE, hwnd: 0,
                                              host_name: client.HOST_NAME,
                                              file_extensions: extensions };
            case "open":             return { hwnd: 0, root_base_type: client.ROOT_BASE_TYPE,
                                              file_extensions: extensions,
                                              stage_assembly: client.STAGE_ASSEMBLY };
            case "search":           return { root_base_type: client.ROOT_BASE_TYPE, hwnd: 0 };

            case "save":             return { item_id: itemId, file_path: filePath };
            //: Save As shows the same New Object form as New, so it declares the same things:
            //: the host's name for the template chip's wording, and what this host can open so a
            //: type whose template it cannot open is not offered. Word sends both; without them
            //: the form offered every type and the chip read "Has template".
            case "save_as_new":      return { file_path: filePath, hwnd: 0, attributes: {},
                                              root_base_type: client.ROOT_BASE_TYPE,
                                              host_name: client.HOST_NAME,
                                              file_extensions: extensions };
            case "save_as_existing": return { file_path: filePath, hwnd: 0,
                                              root_base_type: client.ROOT_BASE_TYPE,
                                              file_extensions: extensions };

            case "check_out":        return { item_id: itemId };
            case "check_in":         return { item_id: itemId, file_path: filePath,
                                              is_assembly: false, structure: [], joints: [],
                                              saved_unsaved_changes: false };
            case "release":          return { item_id: itemId };
            //: The entries are the document's comments, which only the editor can read; the
            //: runner fills them in. The rest is what the service wants to know about the sender.
            case "markup":           return { item_id: itemId, host: client.HOST_NAME,
                                              file_path: filePath, entries: [] };
            //: A GET: which item, and nothing to post.
            case "apply_markups":    return { item_id: itemId };
            case "revise":           return { item_id: itemId, hwnd: 0 };
            case "change_owner":     return { item_id: itemId, file_path: filePath, hwnd: 0 };

            case "worklist":         return { hwnd: 0 };
            case "new_workflow":     return { item_id: itemId };

            case "properties":       return { item_id: itemId, hwnd: 0 };
            case "edit_values":      return { item_id: itemId, hwnd: 0, document_values: {} };
            case "refresh_values":   return { item_id: itemId };
            case "reload_document":  return { item_id: itemId };

            case "settings":         return {};
            case "about":            return { hwnd: 0, host_name: client.HOST_NAME,
                                              addin_version: client.ADDIN_VERSION };

            //: A GET, and the panel reports it; there is nothing to post.
            case "connection":       return null;

            default:                 return { hwnd: 0 };
        }
    }

    /**
     * Why a command is greyed, for its tooltip — so a disabled button is never a dead end.
     * Null when it is not greyed.
     */
    function disabledBecause(command, context) {
        if (isEnabled(command, context)) { return null; }
        var c = context || {};

        if (command.when === SIGNED_OUT) { return "You are already signed in."; }
        if (!c.signedIn) { return "Sign in to Nexus PLM first."; }

        if (command.when === UNREGISTERED) {
            return "This document is already registered in Nexus PLM. Use Save to update it.";
        }

        var state = c.state;
        var inPlm = !!(state && state.status && state.status !== "unknown");
        if (needsDocument(command) && !inPlm) {
            return "Nexus PLM does not know which item this document is.";
        }
        if (command.when === CHECKED_IN) {
            return state.status === "checked_out"
                ? "Check the document in first."
                : "Not available while the document is " + String(state.status).replace(/_/g, " ") + ".";
        }
        if (command.when === MINE) { return "This document is not checked out to you."; }
        if (command.when === NOT_MINE) {
            return "You have this document checked out - your copy is the newer one. Check in first if you want to go back to what PLM holds.";
        }
        return "Not available for this document.";
    }

    return {
        TAB: TAB,
        ALWAYS: ALWAYS, SIGNED_OUT: SIGNED_OUT, SIGNED_IN: SIGNED_IN,
        SESSION: SESSION, UNREGISTERED: UNREGISTERED,
        DOCUMENT: DOCUMENT, CHECKED_IN: CHECKED_IN, MINE: MINE, NOT_MINE: NOT_MINE,
        COMMANDS: COMMANDS,
        GROUPS: GROUPS,
        bodyFor: bodyFor,
        needsDocument: needsDocument,
        needsSession: needsSession,
        sessionOptional: sessionOptional,
        byId: byId,
        inGroup: inGroup,
        topLevel: topLevel,
        childrenOf: childrenOf,
        isEnabled: isEnabled,
        disabledBecause: disabledBecause
    };
}));
