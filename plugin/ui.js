/*
 * The Navigator's DOM: it draws what `lib/navigator.js` decided, and does nothing else.
 *
 * No rule about the tree, the rows or which buttons are live lives here. If a question can be
 * answered without an editor and without a browser it belongs in `lib/navigator.js`, where the
 * tests reach it. This file exists because a DOM cannot be tested that way, and it is
 * deliberately dull.
 *
 * The four commands the pane runs act on the SELECTED row, not on the open document, so they do
 * not go through the runner's command table — that is keyed to the document this frame is docked
 * to. They call the service the way Word's ribbon does on the pane's behalf, and Check In is the
 * one that still needs the document open here, because the service uploads a path.
 */

(function (window, document) {
    "use strict";

    var nav = window.NexusPlmNavigator;
    var panel = window.NexusPlmPanel;
    var clientModule = window.NexusPlmClient;
    var editor = window.NexusPlmEditor;

    var client = new clientModule.Client();

    //: What the pane is drawn from.
    var folders = [];
    var items = [];
    var expanded = {};
    var folderId = null;
    var selected = null;
    var searching = false;
    var username = null;
    var signedIn = false;
    var connected = false;
    var errorText = null;
    //: The document this frame is docked to, so Check In knows whether the row is open here.
    var openPartNumber = null;

    var searchTimer = null;

    function el(id) { return document.getElementById(id); }

    function text(node, value) { node.textContent = value === null || value === undefined ? "" : String(value); }

    // ── reads ────────────────────────────────────────────────────────────────

    function loadFolders(refresh, done) {
        client.folders(refresh).then(function (answer) {
            signedIn = answer.signed_in !== false;
            username = answer.username || username;
            connected = !!answer.success;
            errorText = answer.success ? null : (answer.error || "The folder tree could not be read.");
            folders = (answer.folders || []).map(nav.folderFromWire);
            draw();
            if (done) { done(); }
        });
    }

    function loadItems(id, refresh) {
        searching = false;
        client.folderItems(id, refresh).then(function (answer) {
            signedIn = answer.signed_in !== false;
            errorText = answer.success ? null : (answer.error || "The folder could not be listed.");
            items = (answer.items || []).map(nav.itemFromWire);
            selected = null;
            draw();
        });
    }

    function runSearch(query) {
        client.lookup(query).then(function (answer) {
            signedIn = answer.signed_in !== false;
            errorText = answer.success ? null : (answer.error || "The search could not be run.");
            items = (answer.items || []).map(nav.itemFromWire);
            selected = null;
            searching = true;
            draw();
        });
    }

    // ── the pane's own commands ──────────────────────────────────────────────

    function say(message, severity) {
        if (message) { client.notify(message, severity || "info"); }
    }

    /**
     * Open the selected row. The service stages its primary dataset and answers a path; the
     * document arrives checked in, because taking the lock stays an explicit act — exactly as it
     * is from the Open dialog, and as Word's pane does it.
     */
    function openSelected() {
        var item = selected;
        if (!item) { return; }
        client.stage(item.plmObjectId).then(function (answer) {
            if (!answer || !answer.success || !answer.file_path) {
                say(answer && answer.error ? answer.error
                                           : item.partNumber + " could not be opened.", "warning");
                return;
            }
            editor.rememberPath(answer.file_path);
            editor.openFile(answer.file_path, function (asked) {
                if (!asked) {
                    say("Nexus PLM staged " + item.partNumber + " at " + answer.file_path +
                        " but ONLYOFFICE could not open it from here.", "warning");
                }
            });
        });
    }

    function checkOutSelected() {
        var item = selected;
        if (!item) { return; }
        client.command("/plm/checkout", { item_id: item.plmObjectId }).then(function (answer) {
            if (!answer || answer.success === false) {
                say(answer && answer.error ? answer.error
                                           : item.partNumber + " could not be checked out.", "warning");
            }
            refreshCurrentList();
        });
    }

    /**
     * Check the selected row back in. Only possible for the document this frame is docked to:
     * the service uploads the working copy, and that is the only one this host has a path for.
     * Word's pane refuses the same case with the same sentence.
     */
    function checkInSelected() {
        var item = selected;
        if (!item) { return; }
        editor.document(function (doc) {
            client.command("/plm/checkin", {
                item_id: item.plmObjectId,
                file_path: doc.path,
                is_assembly: false,
                structure: [],
                joints: [],
                saved_unsaved_changes: false
            }).then(function (answer) {
                if (!answer || answer.success === false) {
                    if (answer && !answer.cancelled && answer.error) { say(answer.error, "warning"); }
                }
                refreshCurrentList();
            });
        });
    }

    function propertiesSelected() {
        var item = selected;
        if (!item) { return; }
        client.command("/plm/properties", { item_id: item.plmObjectId, hwnd: 0 }).then(function (answer) {
            if (!answer || answer.success === false) {
                say(answer && answer.error ? answer.error
                                           : "The Properties panel could not be opened.", "warning");
            }
        });
    }

    function refreshCurrentList() {
        if (searching) { return; }
        if (folderId) { loadItems(folderId, true); }
    }

    // ── drawing ──────────────────────────────────────────────────────────────

    function drawTree() {
        var host = el("np-tree");
        host.textContent = "";

        if (!folders.length) {
            var message = document.createElement("div");
            message.className = "np-message";
            text(message, errorText ? "" : "No folders hold anything this plugin can open.");
            if (!errorText) { host.appendChild(message); }
            return;
        }

        nav.flatten(nav.buildTree(folders), expanded, 0).forEach(function (row) {
            var node = document.createElement("div");
            node.className = "np-row";
            node.setAttribute("role", "treeitem");
            node.setAttribute("aria-selected", row.folder.folderId === folderId ? "true" : "false");
            node.style.paddingLeft = (6 + row.depth * 12) + "px";

            var twisty = document.createElement("button");
            twisty.className = "np-twisty";
            twisty.setAttribute("aria-label", row.open ? "Collapse " + row.folder.name : "Expand " + row.folder.name);
            text(twisty, row.hasChildren ? (row.open ? "▼" : "▶") : "");
            if (row.hasChildren) {
                twisty.addEventListener("click", function (e) {
                    e.stopPropagation();
                    if (expanded[row.folder.folderId]) { delete expanded[row.folder.folderId]; }
                    else { expanded[row.folder.folderId] = true; }
                    draw();
                });
            }
            node.appendChild(twisty);

            var name = document.createElement("span");
            name.className = "np-row-name";
            name.title = row.folder.name;
            text(name, row.folder.name);
            node.appendChild(name);

            if (row.folder.hasBadge && row.folder.objectCount > 0) {
                var count = document.createElement("span");
                count.className = "np-count";
                text(count, row.folder.objectCount);
                node.appendChild(count);
            }

            node.addEventListener("click", function () {
                folderId = row.folder.folderId;
                el("np-query").value = "";
                loadItems(folderId, false);
            });
            host.appendChild(node);
        });
    }

    function drawItems() {
        text(el("np-band"), nav.bandHeading({ items: items, searching: searching, folderId: folderId }));

        var host = el("np-list");
        host.textContent = "";

        if (!items.length && folderId && !searching) {
            var message = document.createElement("div");
            message.className = "np-message";
            text(message, "Nothing here this plugin can open.");
            host.appendChild(message);
            return;
        }

        items.forEach(function (item) {
            var card = document.createElement("div");
            card.className = "np-card";
            card.setAttribute("role", "option");
            card.setAttribute("aria-selected", selected && selected.plmObjectId === item.plmObjectId ? "true" : "false");

            var top = document.createElement("div");
            top.className = "np-card-top";

            var part = document.createElement("span");
            part.className = "np-part";
            part.title = item.partNumber;
            text(part, item.partNumber);
            top.appendChild(part);

            var rev = document.createElement("span");
            rev.className = "np-rev";
            text(rev, item.revision || "—");
            top.appendChild(rev);

            var status = document.createElement("span");
            status.className = "np-status";
            text(status, item.status || "draft");
            top.appendChild(status);
            card.appendChild(top);

            if (item.description) {
                var desc = document.createElement("div");
                desc.className = "np-desc";
                desc.title = item.description;
                text(desc, item.description);
                card.appendChild(desc);
            }

            if (item.checkedOutBy) {
                var mine = nav.isMyLock(item.checkedOutBy, username);
                var lock = document.createElement("span");
                lock.className = "np-lock " + (mine ? "np-lock-mine" : "np-lock-theirs");
                text(lock, mine ? "Checked out to you" : "Checked out to " + item.checkedOutBy);
                card.appendChild(lock);
            }

            card.addEventListener("click", function () { selectItem(item); });
            //: Double-click opens, the gesture the Open dialog already uses.
            card.addEventListener("dblclick", function () { selectItem(item); openSelected(); });
            host.appendChild(card);
        });
    }

    /** Picking a search hit reveals where it lives, as Word's pane does. */
    function selectItem(item) {
        selected = item;
        if (searching && folderId) {
            nav.ancestorsOf(folders, folderId).forEach(function (id) { expanded[id] = true; });
        }
        draw();
    }

    function drawActions() {
        var allowed = nav.actionsFor(selected, username, openPartNumber);
        [["np-open", "open"], ["np-checkout", "checkout"],
         ["np-checkin", "checkin"], ["np-properties", "properties"]].forEach(function (pair) {
            var button = el(pair[0]);
            button.disabled = !allowed[pair[1]];
            var why = allowed[pair[1]] ? null : nav.whyDisabled(pair[1], selected, username, openPartNumber);
            button.title = why || "";
        });
    }

    function draw() {
        var trouble = el("np-error");
        if (!clientModule.hasSecret()) {
            trouble.hidden = false;
            text(trouble, panel.NOT_INSTALLED);
        } else if (!signedIn) {
            trouble.hidden = false;
            text(trouble, "You are not signed in to Nexus PLM. Use Sign In on the Nexus PLM tab.");
        } else if (errorText) {
            trouble.hidden = false;
            text(trouble, errorText);
        } else {
            trouble.hidden = true;
        }

        drawTree();
        drawItems();
        drawActions();

        text(el("np-status"), connected
            ? (username ? "Signed in as " + username : "Signed in")
            : "Not connected to Nexus PLM");
    }

    // ── wiring ───────────────────────────────────────────────────────────────

    function start(editorType) {
        el("np-refresh").addEventListener("click", function () {
            loadFolders(true);
            if (folderId) { loadItems(folderId, true); }
        });

        el("np-query").addEventListener("input", function (e) {
            var query = e.target.value;
            if (searchTimer) { window.clearTimeout(searchTimer); searchTimer = null; }

            if (!nav.isSearchable(query)) {
                if (searching) {
                    searching = false;
                    if (folderId) { loadItems(folderId, false); } else { items = []; draw(); }
                }
                return;
            }
            searchTimer = window.setTimeout(function () { runSearch(query); }, nav.SEARCH_DEBOUNCE_MS);
        });

        el("np-open").addEventListener("click", openSelected);
        el("np-checkout").addEventListener("click", checkOutSelected);
        el("np-checkin").addEventListener("click", checkInSelected);
        el("np-properties").addEventListener("click", propertiesSelected);

        //: The panel is its own frame, so it registers with the tray like any other host.
        client.connect();
        var beat = window.setInterval(function () { client.heartbeat(); }, clientModule.HEARTBEAT_MS);
        window.addEventListener("unload", function () {
            window.clearInterval(beat);
            client.disconnect();
        });

        document.title = panel.title(editorType);
        editor.document(function (doc) {
            openPartNumber = doc.title ? String(doc.title).replace(/\.[^.]+$/, "") : null;
            loadFolders(false);
        });
    }

    window.NexusPlmUi = {
        start: start,
        //: The document changed under us, so the folder's rows may have too.
        documentReady: function () { refreshCurrentList(); }
    };

})(window, document);
