/*
 * Talking to the Nexus PLM Addin Service.
 *
 * The service on localhost:5100 owns the dialogs and does the talking to the PLM server, so the
 * same windows, wording and behaviour appear in ONLYOFFICE, LibreOffice, Word, Excel, PowerPoint
 * and FreeCAD. A host declares what it can do and it travels with the request; the service needs
 * no change for a new host. If this plugin ever seems to need a service change, the question is
 * what it should be declaring instead.
 *
 * ── The one thing that is different here, and is NOT yet proven ──────────────────────────────
 * Every other host is a desktop process. This one is JavaScript in a browser frame, so two
 * things stand between it and the service, and both must be settled before the commands can be
 * built (see CLAUDE.md, "The open question"):
 *
 *   1. CORS. The service answers a desktop client that sends no Origin. A browser sends one and
 *      will refuse the response without `Access-Control-Allow-Origin`.
 *   2. Mixed content. A browser on an https:// page will not call http://localhost:5100 at all.
 *
 * Neither is a reason to design differently — the service is still the right place for the
 * dialogs — but until both are measured, nothing here should be assumed to work.
 */

(function (root, factory) {
    if (typeof module === "object" && module.exports) { module.exports = factory(); }
    else { root.NexusPlmClient = factory(); }
}(typeof self !== "undefined" ? self : this, function () {
    "use strict";

    var DEFAULT_BASE_URL = "http://localhost:5100";

    /** How a host says what it can open. The service keeps no list of hosts. */
    var FILE_EXTENSIONS = [".docx", ".dotx", ".xlsx", ".xltx", ".pptx", ".potx"];

    /** A document has no assembly to load; that panel is the CAD flow bleeding in. */
    var STAGE_ASSEMBLY = false;

    function Client(baseUrl, fetchImpl) {
        this.baseUrl = (baseUrl || DEFAULT_BASE_URL).replace(/\/+$/, "");
        //: Injectable so the tests can drive it without a network and without a browser.
        this._fetch = fetchImpl || (typeof fetch !== "undefined" ? fetch.bind(null) : null);
    }

    /**
     * One call, answering a plain object.
     *
     * A refused call is not an exception: the service answers {"success": false, "error": ...}
     * and every caller shows that. Only a service that cannot be reached rejects, because that is
     * the one failure the user has to fix themselves.
     */
    Client.prototype._call = function (method, path, body) {
        if (!this._fetch) { return Promise.reject(new Error("No fetch available.")); }
        var options = { method: method, headers: { "Content-Type": "application/json" } };
        if (body !== undefined) { options.body = JSON.stringify(body); }

        return this._fetch(this.baseUrl + path, options).then(function (response) {
            return response.text().then(function (raw) {
                if (!raw) { return {}; }
                try { return JSON.parse(raw); }
                catch (e) {
                    return { success: false, error: "The service answered something that is not JSON." };
                }
            });
        }, function () {
            return {
                success: false,
                unreachable: true,
                error: "Nexus PLM is not running. Start the Nexus PLM Addins tray application."
            };
        });
    };

    Client.prototype.health = function () { return this._call("GET", "/api/health"); };

    /** Who is signed in: {"success": bool, "username": ...}. */
    Client.prototype.me = function () { return this._call("GET", "/api/auth/me"); };

    /** What PLM knows about a document, by path or by item. */
    Client.prototype.state = function (query) {
        var q = [];
        if (query && query.filePath) { q.push("file_path=" + encodeURIComponent(query.filePath)); }
        if (query && query.itemId) { q.push("item_id=" + encodeURIComponent(query.itemId)); }
        return this._call("GET", "/plm/state" + (q.length ? "?" + q.join("&") : ""));
    };

    /** Every folder the signed-in user may read, for the navigator. */
    Client.prototype.folders = function () { return this._call("GET", "/plm/navigation/folders"); };

    /** What is in one folder. */
    Client.prototype.folderItems = function (folderId) {
        return this._call("GET", "/plm/navigation/folders/" + encodeURIComponent(folderId) + "/items");
    };

    return {
        Client: Client,
        DEFAULT_BASE_URL: DEFAULT_BASE_URL,
        FILE_EXTENSIONS: FILE_EXTENSIONS,
        STAGE_ASSEMBLY: STAGE_ASSEMBLY
    };
}));
