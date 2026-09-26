/*
 * Where the open document is on disk — worked out, because the editor will not say.
 *
 * ── What was measured (Desktop Editors 9.4.0) ───────────────────────────────────────────────
 * A plugin frame learns the document's NAME: `Asc.plugin.info.documentTitle` is the file name,
 * undocumented but present. It does not learn the path. The desktop shell's file getters exist in
 * the frame but answer for the frame that opened the file, which is not this one; a callCommand
 * script runs in a sandbox with neither the shell nor the editor's api; and the shell's recent-
 * files callback never reaches a plugin frame. All three were tried before this file was written.
 *
 * ── So the path is remembered rather than asked for ─────────────────────────────────────────
 * Every file the service hands to this host — Open, New, Search, Revise, Reload — comes with the
 * path it was staged at, and the service names every staged file by its part number. A plugin
 * that remembers "this name lives here" whenever it opens a file can answer the question later,
 * in whichever tab the file is opened in: every tab's frame shares the plugin's own storage.
 * The staging folder is remembered too, so a file the service staged before this plugin ever
 * ran is still found, by name, where the service puts everything.
 *
 * What this cannot tell apart is two copies of one file name in different folders. That is the
 * limit, and it is stated where it bites: an upload is refused, with the reason, when no path is
 * known for the name.
 *
 * Plain functions over a plain store, so the rule is tested without a browser.
 */

(function (root, factory) {
    if (typeof module === "object" && module.exports) { module.exports = factory(); }
    else { root.NexusPlmPaths = factory(); }
}(typeof self !== "undefined" ? self : this, function () {
    "use strict";

    /** The one key in the store. Versioned, so a changed shape is a fresh start, not a crash. */
    var KEY = "nexusplm.paths.v1";

    function baseName(path) {
        var parts = String(path || "").split(/[\\/]/);
        return parts[parts.length - 1] || "";
    }

    function dirName(path) {
        var s = String(path || "");
        var i = Math.max(s.lastIndexOf("\\"), s.lastIndexOf("/"));
        return i > 0 ? s.slice(0, i) : "";
    }

    function join(dir, name) {
        var sep = dir.indexOf("/") !== -1 && dir.indexOf("\\") === -1 ? "/" : "\\";
        return dir + sep + name;
    }

    function read(store) {
        try {
            var raw = store && store.getItem(KEY);
            var data = raw ? JSON.parse(raw) : null;
            if (data && typeof data === "object" && data.files) { return data; }
        } catch (e) { /* unreadable is empty */ }
        return { files: {}, stagingDir: null };
    }

    function write(store, data) {
        try { store.setItem(KEY, JSON.stringify(data)); } catch (e) { /* storage refused; still remembered in memory by the caller */ }
    }

    /** Remember where a file the service staged lives. Names are matched without case. */
    function remember(store, path) {
        var name = baseName(path);
        if (!name) { return; }
        var data = read(store);
        data.files[name.toLowerCase()] = path;
        var dir = dirName(path);
        if (dir) { data.stagingDir = dir; }
        write(store, data);
    }

    /**
     * The path a document of this name was staged at: the exact file if it was opened through
     * this host, else the service's staging folder plus the name if that folder is known, else
     * null.
     */
    function lookup(store, title) {
        var name = baseName(title);
        if (!name) { return null; }
        var data = read(store);
        var exact = data.files[name.toLowerCase()];
        if (exact) { return exact; }
        return data.stagingDir ? join(data.stagingDir, name) : null;
    }

    /** Forget everything, for the tests and for a user starting over. */
    function clear(store) {
        try { store.removeItem(KEY); } catch (e) { /* nothing to forget */ }
    }

    return {
        KEY: KEY,
        baseName: baseName,
        dirName: dirName,
        remember: remember,
        lookup: lookup,
        clear: clear
    };
}));
