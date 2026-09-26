/*
 * The editor, and the desktop application around it, as the plugin can reach them.
 *
 * Everything that touches ONLYOFFICE itself is here and nowhere else: which file is open, whether
 * it is saved, opening another one, saving, reading and writing the document's fields, its
 * comments, and the Navigator panel. Each is one small function, guarded, so that a build of
 * ONLYOFFICE without some capability degrades to "cannot tell" rather than to an exception nobody
 * sees.
 *
 * Two runtimes are used, and they are different things:
 *
 *   window.Asc.plugin         the plugin API: executeMethod for the editor's plugin methods and
 *                             callCommand for a script run in a SANDBOX with the document
 *                             builder's Api. A callCommand function is SERIALISED — it cannot
 *                             close over anything here, and takes its input from Asc.scope.
 *                             MEASURED (Desktop Editors 9.4.0): the sandbox has Api and nothing
 *                             else — no desktop shell, no editor api object.
 *   window.AscDesktopEditor   the DESKTOP shell, injected into every page it hosts, this frame
 *                             included. MEASURED: its per-file getters (LocalFileGetSourcePath,
 *                             LocalFileGetSaved, LocalFileGetOpenChangesCount) answer for the
 *                             frame that opened the file, not this one; LocalFileOpen and
 *                             execCommand("open:recent") are accepted here and open nothing; the
 *                             recent-files callback never arrives here. What DOES work from here
 *                             is the shell's TOOL interface, the one the shipped AI agent plugin
 *                             uses: `callToolFunction("file_opener", {path})` opens a file in a
 *                             new tab, and `callToolFunction("recent_files_reader")` answers
 *                             every recently opened file with its full path — the open ones
 *                             first. `getToolFunctions()` lists them with their schemas.
 *
 * So the document is known by its NAME — `Asc.plugin.info.documentTitle`, undocumented but
 * present — and its path comes from the recent-files tool, matched by that name, with what the
 * service handed over (lib/paths.js) as the fallback. Whether it has unsaved changes cannot be
 * read at all; the builder's Api.Save is asked for before every upload instead, as Word saves
 * before every upload.
 *
 * This file cannot be tested without an editor, and it is deliberately dull: every decision
 * about what to do with what it reads lives in lib/, where the tests are.
 */

(function (window) {
    "use strict";

    var EDITORS = ["word", "cell", "slide"];

    function plugin() { return (window.Asc && window.Asc.plugin) || null; }

    /** Run a plugin method, answering through the callback; a missing runtime answers null. */
    function method(name, args, callback) {
        var p = plugin();
        if (!p || typeof p.executeMethod !== "function") { if (callback) { callback(null); } return; }
        try { p.executeMethod(name, args || [], callback); }
        catch (e) { if (callback) { callback(null); } }
    }

    /** Run a function inside the editor's frame with the builder Api, its input in Asc.scope. */
    function command(fn, scope, callback) {
        var p = plugin();
        if (!p || typeof p.callCommand !== "function") { if (callback) { callback(null); } return; }
        window.Asc.scope = scope || {};
        try { p.callCommand(fn, false, false, callback); }
        catch (e) { if (callback) { callback(null); } }
    }

    // ── the file ─────────────────────────────────────────────────────────────

    var paths = window.NexusPlmPaths;

    function store() {
        try { return window.localStorage; } catch (e) { return null; }
    }

    /** The open file's name, as the editor told this plugin at init. */
    function title() {
        var p = plugin();
        var t = p && p.info && p.info.documentTitle;
        return typeof t === "string" && t ? t : null;
    }

    /** The shell's tool interface: `callToolFunction(name, jsonArgs)` answering a JSON string. */
    function tool(name, args) {
        var d = window.AscDesktopEditor;
        if (!d || typeof d.callToolFunction !== "function") { return null; }
        try {
            var raw = args === undefined ? d.callToolFunction(name) : d.callToolFunction(name, JSON.stringify(args));
            return typeof raw === "string" ? JSON.parse(raw) : raw;
        } catch (e) { return null; }
    }

    /**
     * Where a file of this name is, from the shell's recent-files list: the open documents come
     * first, so the first match is the one this plugin is most likely running in. Two open
     * copies of one name cannot be told apart — the newer wins — and that is the limit.
     */
    function recentPathFor(name) {
        var answer = tool("recent_files_reader");
        var files = answer && answer.files;
        if (!Array.isArray(files) || !name) { return null; }
        var wanted = name.toLowerCase();
        for (var i = 0; i < files.length; i++) {
            var path = files[i] && files[i].path;
            if (typeof path === "string" && paths.baseName(path).toLowerCase() === wanted) { return path; }
        }
        return null;
    }

    /**
     * What is known about the open file: `{ title, path, saved, changes, shell }`, through the
     * callback (kept asynchronous so a source that has to be asked can replace this one).
     *
     * `title` is the file's name. `path` is where it is: the shell's recent-files list first,
     * then wherever the service said it staged a file of that name. `saved` and `changes` are
     * null: the editor does not tell a plugin.
     */
    function document(callback) {
        var t = title();
        callback({
            title: t,
            path: t ? (recentPathFor(t) || paths.lookup(store(), t)) : null,
            saved: null,
            changes: null,
            shell: !!window.AscDesktopEditor
        });
    }

    /** A file the service just handed over: remember where it is, by name, for every tab. */
    function rememberPath(path) {
        if (path) { paths.remember(store(), path); }
    }

    /**
     * Open a file in a new editor tab. Answers, through the callback, whether anything was asked.
     *
     * Called from THIS frame: the shell is here and its LocalFileOpen works from here (measured).
     * The editor's start page reaches the same place through `execCommand("open:recent", ...)`,
     * so that is the fallback.
     */
    function openFile(path, callback) {
        if (!path) { callback(false); return; }
        // The shell's own file opener, as the shipped AI agent uses it. It answers
        // {role, status, error_message}; anything but "success" is a failure to report.
        var answer = tool("file_opener", { path: path });
        if (answer && answer.status === "success") { callback(true); return; }
        callback(false);
    }

    /**
     * Ask the editor to save, the one way a plugin can: the builder's Api.Save, which every one
     * of the three editors provides. Answers whether it could be asked.
     */
    function save(callback) {
        command(function () {
            try { if (typeof Api.Save === "function") { Api.Save(); return true; } } catch (e) { /* no */ }
            return false;
        }, {}, function (asked) { callback(!!asked); });
    }

    /**
     * One line for the Connection Status toast: what this plugin knows about the document, and
     * what the field writer can see in it — so a value that did not land can be explained.
     */
    function probe(callback) {
        document(function (doc) {
            var line = doc.title
                ? "Document: " + doc.title + (doc.path ? ", on disk at " + doc.path + "." : ". Its path is not known: open it through Nexus PLM (Open, Search or New) for uploads.")
                : "No document is open.";
            command(function () {
                var out = [];
                try {
                    var d = Api.GetDocument();
                    var controls = d.GetAllContentControls();
                    var named = 0;
                    for (var i = 0; i < controls.length; i++) {
                        try { if (controls[i].GetAlias() || controls[i].GetTag()) { named++; } } catch (e1) { /* skip */ }
                    }
                    out.push(named + " named content control(s)");
                    var props = null;
                    try { props = d.GetCustomProperties(); } catch (e0) { /* none */ }
                    out.push(props ? "custom properties: yes" : "custom properties: not in this build");
                } catch (e) { out.push("no document builder"); }
                return out.join(", ");
            }, {}, function (seen) { callback(line + " Fields: " + (seen || "not readable") + "."); });
        });
    }

    // ── the document's fields ────────────────────────────────────────────────
    // The same three places the template connectors and the Word add-in use: a document's
    // content controls (by alias, then tag) and custom properties; a spreadsheet's defined
    // names; a presentation's named shapes. Values PLM owns arrive already written into a staged
    // file by the service, so these only matter for Edit Values, Refresh Values and Save As.
    //
    // MEASURED (Desktop Editors 9.4.0): a content control tagged with the attribute name shows
    // the written value at once. A custom property is written and reads back, but a DOCPROPERTY
    // field showing it keeps its cached text - ONLYOFFICE's SDK parses MERGEFIELD, ADDIN,
    // FORMTEXT, PAGE, REF, TOC and the like, and not DOCPROPERTY. So a template meant for this
    // host puts each field in a content control whose tag is the attribute name: the Word add-in
    // writes those too, and the Vault's Word connector reads them by tag, then alias.

    var READERS = {
        word: function () {
            var out = {};
            try {
                var doc = Api.GetDocument();
                var names = Asc.scope.names || [];
                try {
                    var props = doc.GetCustomProperties();
                    for (var n = 0; n < names.length; n++) {
                        try { var v = props.Get(names[n]); if (v !== undefined && v !== null) { out[names[n]] = String(v); } } catch (e1) { /* absent */ }
                    }
                } catch (e2) { /* no custom properties API */ }
                var controls = doc.GetAllContentControls();
                for (var i = 0; i < controls.length; i++) {
                    try {
                        var c = controls[i];
                        var key = c.GetAlias() || c.GetTag();
                        if (!key) { continue; }
                        var text = "";
                        try { text = c.GetRange().GetText(); }
                        catch (e3) { try { text = c.GetContent().GetRange().GetText(); } catch (e4) { /* empty */ } }
                        out[key] = text || "";
                    } catch (e5) { /* locked or odd control */ }
                }
            } catch (e) { /* no document */ }
            return out;
        },
        cell: function () {
            var out = {};
            try {
                var names = Asc.scope.names || [];
                for (var n = 0; n < names.length; n++) {
                    try {
                        var dn = Api.GetDefName(names[n]);
                        if (!dn) { continue; }
                        var v = dn.GetRefersToRange().GetValue();
                        if (v !== undefined && v !== null) { out[names[n]] = String(v); }
                    } catch (e1) { /* not a name */ }
                }
            } catch (e) { /* no workbook */ }
            return out;
        },
        slide: function () {
            var out = {};
            try {
                var names = Asc.scope.names || [];
                var pres = Api.GetPresentation();
                var count = pres.GetSlidesCount();
                for (var s = 0; s < count; s++) {
                    var shapes = pres.GetSlideByIndex(s).GetAllShapes();
                    for (var i = 0; i < shapes.length; i++) {
                        try {
                            var shape = shapes[i];
                            var name = typeof shape.GetName === "function" ? shape.GetName() : null;
                            if (!name || names.indexOf(name) === -1) { continue; }
                            out[name] = shape.GetDocContent().GetRange().GetText() || "";
                        } catch (e1) { /* not a text shape */ }
                    }
                }
            } catch (e) { /* no presentation */ }
            return out;
        }
    };

    var WRITERS = {
        word: function () {
            var values = Asc.scope.values || {};
            var written = 0;
            try {
                var doc = Api.GetDocument();
                try {
                    var props = doc.GetCustomProperties();
                    for (var name in values) {
                        if (!values.hasOwnProperty(name) || values[name] === null || values[name] === undefined) { continue; }
                        try { props.Add(name, String(values[name])); written++; } catch (e1) { /* refused */ }
                    }
                } catch (e2) { /* no custom properties API */ }
                // NOT UpdateAllFields: ONLYOFFICE does not parse DOCPROPERTY, and updating turns
                // such a field into "Error! Reference source not found" (measured). A property
                // shown through a DOCPROPERTY field keeps its cached text here; a template meant
                // for ONLYOFFICE shows the value through a content control instead - see below.
                var controls = doc.GetAllContentControls();
                for (var i = 0; i < controls.length; i++) {
                    try {
                        var c = controls[i];
                        var key = c.GetAlias() || c.GetTag();
                        if (!key || !values.hasOwnProperty(key)) { continue; }
                        var v = values[key];
                        // Word leaves a field alone for an empty value rather than blanking it.
                        if (v === null || v === undefined || v === "") { continue; }
                        if (c.GetClassType() === "inlineLvlSdt") {
                            c.RemoveAllElements();
                            c.AddText(String(v));
                        } else {
                            var content = c.GetContent();
                            content.RemoveAllElements();
                            var p = Api.CreateParagraph();
                            p.AddText(String(v));
                            content.Push(p);
                        }
                        written++;
                    } catch (e3) { /* locked control */ }
                }
            } catch (e) { /* no document */ }
            return written;
        },
        cell: function () {
            var values = Asc.scope.values || {};
            var written = 0;
            try {
                for (var name in values) {
                    if (!values.hasOwnProperty(name)) { continue; }
                    var v = values[name];
                    if (v === null || v === undefined || v === "") { continue; }
                    try {
                        var dn = Api.GetDefName(name);
                        if (!dn) { continue; }
                        dn.GetRefersToRange().SetValue(String(v));
                        written++;
                    } catch (e1) { /* not a name */ }
                }
            } catch (e) { /* no workbook */ }
            return written;
        },
        slide: function () {
            var values = Asc.scope.values || {};
            var written = 0;
            try {
                var pres = Api.GetPresentation();
                var count = pres.GetSlidesCount();
                for (var s = 0; s < count; s++) {
                    var shapes = pres.GetSlideByIndex(s).GetAllShapes();
                    for (var i = 0; i < shapes.length; i++) {
                        try {
                            var shape = shapes[i];
                            var name = typeof shape.GetName === "function" ? shape.GetName() : null;
                            if (!name || !values.hasOwnProperty(name)) { continue; }
                            var v = values[name];
                            if (v === null || v === undefined || v === "") { continue; }
                            var content = shape.GetDocContent();
                            content.RemoveAllElements();
                            var p = Api.CreateParagraph();
                            p.AddText(String(v));
                            content.Push(p);
                            written++;
                        } catch (e1) { /* not a text shape */ }
                    }
                }
            } catch (e) { /* no presentation */ }
            return written;
        }
    };

    /** The document's field values, keyed by field name. `names` says which properties to ask for. */
    function readFields(editorType, names, callback) {
        var reader = READERS[editorType];
        if (!reader) { callback({}); return; }
        command(reader, { names: names || [] }, function (out) { callback(out || {}); });
    }

    /** Write PLM's values into the document's fields. Answers how many were written. */
    function writeFields(editorType, values, callback) {
        var writer = WRITERS[editorType];
        if (!writer || !values) { callback(0); return; }
        command(writer, { values: values }, function (n) { callback(typeof n === "number" ? n : 0); });
    }

    // ── comments ─────────────────────────────────────────────────────────────

    /** Every comment in the document, as `[{Id, Data}]`. */
    function comments(callback) {
        method("GetAllComments", [], function (list) { callback(Array.isArray(list) ? list : []); });
    }

    /** Add one comment; `data` is the shape GetAllComments answers with. */
    function addComment(data, callback) {
        method("AddComment", [data], function () { if (callback) { callback(); } });
    }

    // ── the Navigator panel ──────────────────────────────────────────────────

    var panel = null;

    /** Show the Navigator beside the document, or hide it if it is showing. */
    function togglePanel() {
        if (panel) {
            try { panel.close(); } catch (e) { /* already gone */ }
            panel = null;
            return false;
        }
        if (!window.Asc || typeof window.Asc.PluginWindow !== "function") { return null; }
        panel = new window.Asc.PluginWindow();
        panel.show({
            url: "index.html",
            description: "Navigator",
            isVisual: true,
            isModal: false,
            isCanDocked: true,
            dockedPlace: "panel",
            type: "panel",
            EditorsSupport: EDITORS,
            size: [320, 500],
            buttons: [],
            icons: "icons/%theme-type%(light|dark)/navigator%scale%(default).png"
        });
        return true;
    }

    /** ONLYOFFICE reports a window closing through Asc.plugin.button(-1, windowId). */
    function windowClosed(windowId) {
        if (panel && panel.id === windowId) { panel = null; }
    }

    function panelOpen() { return !!panel; }

    window.NexusPlmEditor = {
        document: document,
        rememberPath: rememberPath,
        openFile: openFile,
        save: save,
        readFields: readFields,
        writeFields: writeFields,
        comments: comments,
        addComment: addComment,
        togglePanel: togglePanel,
        windowClosed: windowClosed,
        panelOpen: panelOpen,
        probe: probe
    };

})(window);
