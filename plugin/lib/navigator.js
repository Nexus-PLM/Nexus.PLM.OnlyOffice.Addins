/*
 * The docked Navigator: the vault as a folder tree, and what may be done to the row you pick.
 *
 * This is the same pane Word docks beside the document (`src/embed` in the React client, shipped
 * into Word through a WebView2). Everything it decides that does not need an editor is here, so
 * the tree, the wire mapping and the action rules are tested without one — and so that the two
 * hosts answer the same questions the same way.
 *
 * ── What is deliberately the same as Word's ─────────────────────────────────────────────────
 * The four reads (`/plm/navigation/folders`, `.../folders/{id}/items`, `.../lookup`, and
 * `/plm/navigation/stage` to open one), the tree's shape and sibling order, the four commands
 * Open / Check Out / Check In / Properties, and the rules that grey them.
 *
 * ── What differs, and why ───────────────────────────────────────────────────────────────────
 * Word's pane cannot run a command: a WebView2 has no window handle and does not know the open
 * document's path, so it posts a message and the ribbon runs it. A plugin frame here IS the
 * host, so the panel calls the service itself. The one thing it still cannot do is check in a
 * document it does not have open, and Word cannot either — the same sentence is used.
 */

(function (root, factory) {
    if (typeof module === "object" && module.exports) { module.exports = factory(); }
    else { root.NexusPlmNavigator = factory(); }
}(typeof self !== "undefined" ? self : this, function () {
    "use strict";

    /** Below this a query is too broad to be worth a round trip. Word's number. */
    var MIN_QUERY_LENGTH = 2;

    /** How long after the last keystroke to search. Word's number. */
    var SEARCH_DEBOUNCE_MS = 300;

    // ── the wire ─────────────────────────────────────────────────────────────
    // The service answers snake_case (its JSON policy is SnakeCaseLower). Reading it in one place
    // keeps every consumer from having to know that.

    /** One folder of the tree. `objectCount` counts what THIS host can open, not what is filed. */
    function folderFromWire(w) {
        w = w || {};
        return {
            folderId: w.folder_id || "",
            name: w.name || "",
            parentId: w.parent_id === undefined ? null : w.parent_id,
            isSystem: !!w.is_system,
            sortOrder: w.sort_order || 0,
            objectCount: w.object_count || 0,
            hasBadge: !!w.has_badge
        };
    }

    /** One document row, at its latest revision. */
    function itemFromWire(w) {
        w = w || {};
        return {
            plmObjectId: w.plm_object_id || "",
            revisionId: w.revision_id || "",
            partNumber: w.part_number || "",
            typeName: w.type_name || "",
            description: w.description === undefined ? null : w.description,
            revision: w.revision || "",
            status: w.status || "",
            modifiedAt: w.modified_at || "",
            checkedOutBy: w.checked_out_by === undefined ? null : w.checked_out_by
        };
    }

    // ── the tree ─────────────────────────────────────────────────────────────

    /**
     * Builds the tree from the flat folder list.
     *
     * The service keeps every ancestor of a folder worth showing, so the list arrives connected —
     * but a folder whose parent is missing is promoted to the root rather than dropped. Dropping
     * it is how a tree silently loses a branch, and a pane that is a user's only view of the
     * vault cannot afford to lose one to a data oddity. (Word's `buildTree`, same rule.)
     *
     * Siblings sort by `sortOrder`, then by name, so this matches the web client and the Open
     * dialog.
     */
    function buildTree(folders) {
        var nodes = {};
        var order = [];
        (folders || []).forEach(function (f) {
            if (!f || !f.folderId) { return; }
            nodes[f.folderId] = { folder: f, children: [] };
            order.push(f.folderId);
        });

        var roots = [];
        order.forEach(function (id) {
            var node = nodes[id];
            var parent = node.folder.parentId === null ? null : nodes[node.folder.parentId];
            if (parent && parent !== node) { parent.children.push(node); }
            else { roots.push(node); }
        });

        function bySibling(a, b) {
            return (a.folder.sortOrder - b.folder.sortOrder) ||
                   String(a.folder.name).localeCompare(String(b.folder.name));
        }
        function sortDeep(list) {
            list.sort(bySibling);
            list.forEach(function (n) { sortDeep(n.children); });
            return list;
        }
        return sortDeep(roots);
    }

    /**
     * The ids to expand so `folderId` is on screen — every ancestor, the folder itself excluded.
     * Used when the search box finds a document and the pane reveals where it lives.
     */
    function ancestorsOf(folders, folderId) {
        var parentOf = {};
        (folders || []).forEach(function (f) { if (f && f.folderId) { parentOf[f.folderId] = f.parentId; } });

        var path = [];
        var seen = {};
        seen[folderId] = true;
        var cur = parentOf[folderId] === undefined ? null : parentOf[folderId];
        while (cur !== null && cur !== undefined && !seen[cur]) {
            path.push(cur);
            seen[cur] = true;
            cur = parentOf[cur] === undefined ? null : parentOf[cur];
        }
        return path;
    }

    /** Every node of a tree, depth first, each with the depth it is drawn at. */
    function flatten(nodes, expanded, depth) {
        var out = [];
        (nodes || []).forEach(function (node) {
            var open = !!(expanded && expanded[node.folder.folderId]);
            out.push({ folder: node.folder, depth: depth || 0, open: open, hasChildren: node.children.length > 0 });
            if (open && node.children.length) {
                out = out.concat(flatten(node.children, expanded, (depth || 0) + 1));
            }
        });
        return out;
    }

    // ── locks and the actions ────────────────────────────────────────────────

    /** Whose check-out a lock is. Case and padding do not decide. (Word's `isMyLock`.) */
    function isMyLock(holder, username) {
        if (!holder || !username) { return false; }
        return String(holder).trim().toLowerCase() === String(username).trim().toLowerCase();
    }

    /**
     * Which of the pane's four commands may be pressed for the selected row.
     *
     * Word's pane: Open needs a selection; Check Out needs one nobody holds; Check In needs your
     * own lock; Properties needs a selection. `openHere` is this host's one addition — the
     * document the panel is docked to — because checking in uploads the working copy and only a
     * document open HERE has one.
     */
    function actionsFor(item, username, openPartNumber) {
        if (!item) {
            return { open: false, checkout: false, checkin: false, properties: false };
        }
        var mine = isMyLock(item.checkedOutBy, username);
        var isOpenHere = !!openPartNumber &&
            String(openPartNumber).trim().toLowerCase() === String(item.partNumber).trim().toLowerCase();
        return {
            open: true,
            checkout: !item.checkedOutBy,
            checkin: mine && isOpenHere,
            properties: true
        };
    }

    /** Why a greyed action is greyed, so no button in the pane is a dead end. */
    function whyDisabled(action, item, username, openPartNumber) {
        if (!item) { return "Pick a document first."; }
        var mine = isMyLock(item.checkedOutBy, username);
        if (action === "checkout" && item.checkedOutBy) {
            return mine ? "You already have this checked out."
                        : "Checked out to " + item.checkedOutBy + ".";
        }
        if (action === "checkin") {
            if (!mine) {
                return item.checkedOutBy ? "Checked out to " + item.checkedOutBy + "."
                                         : "This document is not checked out to you.";
            }
            // Word's sentence, for the same reason: the service uploads a path, and only the
            // document this panel is docked to has one here.
            return item.partNumber + " is checked out to you but is not open here. " +
                   "Open it first, then check it in.";
        }
        return null;
    }

    /** What the item band's heading says. */
    function bandHeading(state) {
        var s = state || {};
        var n = (s.items || []).length;
        if (s.searching) { return n + (n === 1 ? " match" : " matches"); }
        if (!s.folderId) { return "Select a folder"; }
        return n + (n === 1 ? " document" : " documents");
    }

    /** Whether what the user typed is worth searching for. */
    function isSearchable(query) {
        return String(query || "").trim().length >= MIN_QUERY_LENGTH;
    }

    return {
        MIN_QUERY_LENGTH: MIN_QUERY_LENGTH,
        SEARCH_DEBOUNCE_MS: SEARCH_DEBOUNCE_MS,
        folderFromWire: folderFromWire,
        itemFromWire: itemFromWire,
        buildTree: buildTree,
        ancestorsOf: ancestorsOf,
        flatten: flatten,
        isMyLock: isMyLock,
        actionsFor: actionsFor,
        whyDisabled: whyDisabled,
        bandHeading: bandHeading,
        isSearchable: isSearchable
    };
}));
