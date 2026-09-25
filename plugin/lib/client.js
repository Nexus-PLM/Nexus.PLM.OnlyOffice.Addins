/*
 * Talking to the Nexus PLM Addin Service.
 *
 * The service on localhost:5100 owns the dialogs and does the talking to the PLM server, so the
 * same windows, wording and behaviour appear in ONLYOFFICE, LibreOffice, Word, Excel, PowerPoint
 * and FreeCAD. A host declares what it can do and it travels with the request; the service needs
 * no change for a new host. If this plugin ever seems to need a service change, the question is
 * what it should be declaring instead.
 *
 * ── Proving this call is local ───────────────────────────────────────────────────────────────
 * Every other host is a desktop process and sends no Origin at all. This one is a page loaded
 * from disk, so it has an OPAQUE origin, which a browser reports as the literal `Origin: null`.
 * The service allows that origin — it has to, or no answer here could ever be read — but an
 * opaque origin proves nothing: any website can obtain one by putting its script in a sandboxed
 * iframe, and one opaque origin cannot be told from another.
 *
 * So a caller with that origin must also present the per-install secret, which the service writes
 * to a file on this machine. A page loaded from disk can read a file beside itself; a remote site
 * cannot read the user's disk. That asymmetry is the whole point, and it is why the secret
 * arrives as `secret.js` loaded by a <script> tag rather than fetched: a browser refuses `fetch`
 * from a file:// page to a file:// URL, but it will happily load a sibling script.
 *
 * Without `secret.js` the calls answer 403 and the panel says so. That is the correct behaviour
 * for a plugin somebody copied into place by hand without running the install step.
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

    /** The header the service reads the per-install secret from. */
    var SECRET_HEADER = "X-Nexus-Addin";

    function Client(baseUrl, fetchImpl, secret) {
        this.baseUrl = (baseUrl || DEFAULT_BASE_URL).replace(/\/+$/, "");
        //: Injectable so the tests can drive it without a network and without a browser.
        this._fetch = fetchImpl || (typeof fetch !== "undefined" ? fetch.bind(null) : null);
        //: Written by the install step into secret.js, which index.html loads before this file.
        this._secret = secret !== undefined
            ? secret
            : (typeof self !== "undefined" ? self.NexusPlmSecret : undefined);
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
        if (this._secret) { options.headers[SECRET_HEADER] = this._secret; }
        if (body !== undefined) { options.body = JSON.stringify(body); }

        return this._fetch(this.baseUrl + path, options).then(function (response) {
            return response.text().then(function (raw) {
                var answer;
                if (!raw) { answer = {}; }
                else {
                    try { answer = JSON.parse(raw); }
                    catch (e) {
                        answer = { success: false,
                                   error: "The service answered something that is not JSON." };
                    }
                }
                // The status travels with the answer, under a name of OUR choosing. Not `status`:
                // a /plm/state answer already has one, and overwriting it put the HTTP 400 from a
                // stateless call straight into the panel's Status row - and made a document that
                // PLM had never heard of look like one it knew.
                answer.httpStatus = response.status;
                answer.refused = response.status === 403;
                return answer;
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

    /** Whether this plugin has been given the secret it needs to be trusted as local. */
    function hasSecret() {
        return !!(typeof self !== "undefined" && self.NexusPlmSecret);
    }

    return {
        Client: Client,
        hasSecret: hasSecret,
        SECRET_HEADER: SECRET_HEADER,
        DEFAULT_BASE_URL: DEFAULT_BASE_URL,
        FILE_EXTENSIONS: FILE_EXTENSIONS,
        STAGE_ASSEMBLY: STAGE_ASSEMBLY
    };
}));
