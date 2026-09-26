/*
 * The host's half of every command — what ONLYOFFICE itself has to do around a service call.
 *
 * The service owns the dialogs and the talking to PLM, and every command answers the same way:
 * `{success, error, cancelled, ...}` plus whatever that command produces. But some of what it
 * produces is for the HOST to act on: a file it staged that must now be opened, an item the open
 * document has just become, values that belong in the document's fields. In Word that half is
 * spread through twenty ribbon handlers. Here it is one table of decisions, taken from those
 * handlers and tested without an editor, so that what a command does after the service answers
 * cannot depend on which button it was pressed from.
 *
 * Nothing in here touches ONLYOFFICE. It answers "what should happen" as plain data — effects —
 * and `background.js` carries them out.
 *
 * ── What this host cannot do that Word can ─────────────────────────────────────────────────
 * Word saves the document before uploading it. ONLYOFFICE has no plugin method that saves a
 * local file: its save path (`LocalFileSave` in sdk-all.js) is the editor's own, driven by the
 * user's Save, and nothing in the plugin API list reaches it. So a command that uploads the
 * file is refused, with the reason, while the document has unsaved changes — measured through
 * `AscDesktopEditor.LocalFileGetOpenChangesCount()`, which is what the editor's own title bar
 * asterisk reads. Saying "save first" is honest; uploading a stale file is not.
 */

(function (root, factory) {
    if (typeof module === "object" && module.exports) {
        module.exports = factory(require("./commands.js"));
    } else {
        root.NexusPlmHost = factory(root.NexusPlmCommands);
    }
}(typeof self !== "undefined" ? self : this, function (commands) {
    "use strict";

    // ── the effects ──────────────────────────────────────────────────────────
    /** Open a file the service staged, in a new editor tab. `path` is where it is. */
    var OPEN = "open";
    /** The open document is now this item: remember it, so the next command knows. */
    var BIND = "bind";
    /** Write PLM's values into the document's fields. `values` is template field to value. */
    var VALUES = "values";
    /** Tell the user something, through the tray's toast. */
    var SAY = "say";

    var CANCEL_WORDS = /cancel/i;

    function say(message, severity) { return { type: SAY, message: message, severity: severity || "info" }; }
    function open(path) { return { type: OPEN, path: path }; }
    function values(mappings) { return { type: VALUES, values: mappings }; }
    function bind(itemId, status, owner, revision) {
        return { type: BIND, itemId: itemId, status: status, owner: owner || null, revision: revision || null };
    }

    /** Whether the answer says the user closed the dialog — their own decision, and no news. */
    function wasCancelled(answer) {
        return !!(answer && (answer.cancelled || CANCEL_WORDS.test(answer.error || "")));
    }

    /** The one toast for an unanswered command, as Word's `ReportUnanswered` words it. */
    function unanswered(answer, fallback) {
        if (answer && answer.unreachable) { return say(answer.error, "error"); }
        // The service announces its own refusals; only a silence is the host's to report.
        if (answer && answer.error) { return null; }
        // A body the service could not bind answers HTTP 400 with a problem report and no
        // toast. That is a bug in this plugin, and it should say which field.
        if (answer && answer.httpStatus === 400) {
            var fields = answer.errors ? Object.keys(answer.errors).join(", ") : "";
            return say("Nexus PLM rejected the request" + (fields ? " (" + fields + ")" : "") +
                       ": " + (answer.title || "bad request") + ".", "error");
        }
        return say(fallback, "error");
    }

    /**
     * What has to be true of the document before this command may run at all.
     *
     * `doc` is what is known about the open file: `{ title, path, saved, changes }` — its name,
     * where it is on disk, whether it has been saved at all, and how many unsaved changes it
     * carries. An upload needs the path, so no path is a refusal that says how to get one. Null
     * for the rest means the check cannot be made, and the command goes ahead: refusing on a
     * guess is worse than a refusal the service can explain.
     *
     * Returns the reason to refuse, or null.
     */
    function refuse(command, doc) {
        var d = doc || {};
        if (!command.saves) { return null; }

        if (!d.path) {
            return "Nexus PLM cannot tell where this file is on disk. Open it through Nexus PLM " +
                   "(Open, Search or New), then " + command.label + ".";
        }

        if (d.saved === false) {
            return "Save the document to a file first, then " + command.label + ".";
        }
        if (typeof d.changes === "number" && d.changes > 0) {
            return "The document has unsaved changes. Save it in ONLYOFFICE (Ctrl+S), then " +
                   command.label + ".";
        }
        return null;
    }

    /**
     * What the host does with a command's answer.
     *
     * `context` is `{ user }` — who is signed in, which is who a fresh lock belongs to.
     */
    function effectsFor(command, answer, context) {
        var a = answer || {};
        var user = (context && context.user) || null;
        var out = [];

        if (!a.success) {
            if (wasCancelled(a)) { return out; }
            var report = unanswered(a, "No response from Nexus PLM for " + command.label + ".");
            if (report) { out.push(report); }
            return out;
        }

        // ── A staged file is opened in a NEW editor tab ─────────────────────────────────────
        // Word opens it in the same application and binds it there. Here every editor tab runs
        // its own copy of this plugin, so the new tab works out which item its file is from the
        // file's own name (the service names every staged file by its part number, and
        // /plm/state resolves a path by that). Nothing about the newly opened file is therefore
        // remembered HERE: this frame belongs to the document that was open when the button was
        // pressed, and binding it to the other document is how a Save would go to the wrong item.
        switch (command.id) {
            case "new":
                if (!a.plm_object_id) { out.push(say("Nexus PLM reported no object for the new document.", "warning")); break; }
                if (a.file_path) { out.push(open(a.file_path)); }
                else { out.push(say("Nexus PLM staged no file for " + (a.part_number || "the new document") + ".", "warning")); }
                break;

            case "open":
                if (!a.file_path) { out.push(say("No file available in the vault for this document.", "warning")); break; }
                out.push(open(a.file_path));
                break;

            case "search":
                if (!a.plm_object_id) { break; }
                if (a.file_path) {
                    out.push(open(a.file_path));
                } else {
                    out.push(say("Selected: " + (a.part_number || a.plm_object_id) +
                                 (a.type_name ? " (" + a.type_name + ")" : ""), "info"));
                }
                break;

            case "save_as_new":
                if (!a.plm_object_id) { break; }
                // Registering a document checks it out to whoever registered it - stamp what the
                // service says it granted, not what was hoped for.
                out.push(bind(a.plm_object_id, a.checked_out ? "checked_out" : "checked_in", a.checked_out ? user : null));
                if (a.attribute_mappings) { out.push(values(a.attribute_mappings)); }
                break;

            case "save_as_existing":
                if (!a.item_id) { break; }
                out.push(bind(a.item_id, a.checked_out ? "checked_out" : "checked_in", a.checked_out ? user : null, a.revision));
                break;

            case "check_in":
                if (a.item_id) { out.push(bind(a.item_id, "checked_in", null, a.revision)); }
                break;

            case "release":
                out.push(bind(null, "released", null, a.revision));
                break;

            case "revise":
                // With a staged file, the new revision opens beside this document and is its own
                // document from then on. Without one, the document in front of the user IS the
                // new revision now, and revising checked it out to them - marking it superseded
                // greyed the very commands they revised in order to use.
                if (a.file_path) {
                    out.push(open(a.file_path));
                    out.push(say("Revision " + (a.revision || "") + " opened in a new tab. This tab still shows the revision it was.", "info"));
                    break;
                }
                if (a.item_id) {
                    out.push(bind(a.item_id, a.checked_out ? "checked_out" : "checked_in", a.checked_out ? user : null, a.revision));
                }
                if (a.attribute_mappings) { out.push(values(a.attribute_mappings)); }
                break;

            case "edit_values":
                // A save can move what the document shows, so the service hands back what it
                // should now display. Nothing saved means nothing to write.
                if (a.saved && a.attribute_mappings) { out.push(values(a.attribute_mappings)); }
                break;

            case "refresh_values":
                if (a.attribute_mappings) { out.push(values(a.attribute_mappings)); }
                break;

            case "reload_document":
                // Word closes the document and reopens what came down. A plugin cannot close the
                // document it lives in, so the current version opens beside it and says so.
                if (!a.file_path) { out.push(say("Could not download the document from Nexus PLM.", "error")); break; }
                out.push(open(a.file_path));
                out.push(say("The version in Nexus PLM has been opened in a new tab. Close this one without saving to discard your copy.", "info"));
                break;

            case "settings":
                if (a.restart_required) {
                    out.push(say("Settings saved. Changes to the service port or log level take effect when Nexus PLM is restarted.", "info"));
                }
                break;

            default:
                // check_out, change_owner, worklist, new_workflow, properties, about, sign in and
                // out, markup: the service raised the toast, and the runner re-reads the state.
                break;
        }
        return out;
    }

    /**
     * Which commands go back to the service for the document's state afterwards.
     *
     * Everything that can move the lock, the revision or the lifecycle. Reading it back is one
     * GET, and it is how the tab tells the truth after a Check In rather than after the next
     * poll.
     */
    function movesState(command) {
        return ["sign_in", "sign_out", "new", "open", "search", "save_as_new", "save_as_existing",
                "check_out", "check_in", "release", "revise", "change_owner", "reload_document",
                "new_workflow"].indexOf(command.id) !== -1;
    }

    return {
        OPEN: OPEN, BIND: BIND, VALUES: VALUES, SAY: SAY,
        refuse: refuse,
        effectsFor: effectsFor,
        wasCancelled: wasCancelled,
        movesState: movesState
    };
}));
