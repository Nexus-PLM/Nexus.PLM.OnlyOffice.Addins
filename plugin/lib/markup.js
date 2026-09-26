/*
 * Review markups: the editor's comments as PLM records them, and PLM's markups as comments.
 *
 * Word sends its comments and tracked changes to `/plm/markup` as entries and brings every
 * reviewer's markups back in as comments. ONLYOFFICE exposes its comments to a plugin through
 * `GetAllComments` and `AddComment`, each as `{Id, Data}` where Data is
 * `{Text, Time, UserName, QuoteText, Solved, UserData, Replies}` — read from the SDK's own
 * serializer (`oTc`/`dze` in sdk-all.js 9.4.0) rather than from the docs. Tracked changes have
 * no plugin method that reads them, so only comments travel.
 *
 * Plain data in, plain data out, so the tests can hold every rule without an editor.
 */

(function (root, factory) {
    if (typeof module === "object" && module.exports) { module.exports = factory(); }
    else { root.NexusPlmMarkup = factory(); }
}(typeof self !== "undefined" ? self : this, function () {
    "use strict";

    var COMMENT = "Comment";
    var OPEN = "Open";
    var RESOLVED = "Resolved";

    /**
     * How a comment this plugin wrote is recognised later. `UserData` is the editor's free field
     * on a comment, and it survives save and reload; the applied entry's id goes there so a
     * second Apply does not add the same markup twice, and Markup does not send PLM's own
     * markups back to it.
     */
    var APPLIED_PREFIX = "nexus-plm:";

    /** ONLYOFFICE keeps a comment's time as milliseconds since the epoch, in a string. */
    function whenOf(time) {
        var n = typeof time === "number" ? time : parseInt(time, 10);
        var d = isNaN(n) ? new Date() : new Date(n);
        return d.toISOString();
    }

    function appliedIdOf(data) {
        var u = data && data.UserData;
        return (typeof u === "string" && u.indexOf(APPLIED_PREFIX) === 0) ? u.slice(APPLIED_PREFIX.length) : null;
    }

    /**
     * The editor's comments as the entries `/plm/markup` takes.
     *
     * Every comment is one entry, keyed by the editor's own stable comment id so a re-send
     * updates rather than duplicates; each reply is an entry of its own under it. A comment
     * that came FROM PLM is not sent back.
     */
    function toEntries(comments) {
        var entries = [];
        (comments || []).forEach(function (comment, ordinal) {
            var data = (comment && comment.Data) || {};
            if (!comment || !comment.Id || appliedIdOf(data)) { return; }

            entries.push({
                external_id: String(comment.Id),
                parent_id: null,
                kind: COMMENT,
                author: data.UserName || "",
                created_at: whenOf(data.Time),
                text: data.Text || "",
                quote: data.QuoteText || null,
                ordinal: ordinal,
                status: data.Solved ? RESOLVED : OPEN
            });

            (data.Replies || []).forEach(function (reply, i) {
                entries.push({
                    external_id: String(comment.Id) + "/" + i,
                    parent_id: String(comment.Id),
                    kind: COMMENT,
                    author: reply.UserName || "",
                    created_at: whenOf(reply.Time),
                    text: reply.Text || "",
                    quote: null,
                    ordinal: i,
                    status: OPEN
                });
            });
        });
        return entries;
    }

    /**
     * Which of PLM's markups to add to the document, as `AddComment` payloads.
     *
     * Everything every reviewer left, except what is already here: an entry this document
     * produced (its external id is one of the comment ids) and an entry a previous Apply
     * already wrote (its id is on the comment). Replies come with their comment. The author is
     * named as PLM knows them, with the host that captured the markup, so a comment from a
     * FreeCAD review does not look like it was typed here.
     */
    function planApply(markups, comments) {
        var present = {};
        (comments || []).forEach(function (c) {
            if (!c || !c.Id) { return; }
            present[String(c.Id)] = true;
            var applied = appliedIdOf(c.Data);
            if (applied) { present[applied] = true; }
        });

        var additions = [];
        (markups || []).forEach(function (markup) {
            var entries = (markup && markup.entries) || [];
            var byId = {};
            entries.forEach(function (e) { if (e && e.id) { byId[e.id] = e; } });

            entries.forEach(function (entry) {
                if (!entry || entry.parent_id) { return; }
                if (present[entry.external_id] || (entry.id && present[entry.id])) { return; }

                var replies = entries.filter(function (e) { return e && e.parent_id === entry.id; });
                additions.push({
                    Text: entry.text || "",
                    Time: String(Date.parse(entry.created_at) || Date.now()),
                    UserName: author(entry, markup),
                    QuoteText: entry.quote || "",
                    Solved: entry.status === RESOLVED,
                    UserData: APPLIED_PREFIX + (entry.id || entry.external_id),
                    Replies: replies.map(function (r) {
                        return { Text: r.text || "", Time: String(Date.parse(r.created_at) || Date.now()),
                                 UserName: author(r, markup) };
                    })
                });
            });
        });
        return additions;
    }

    function author(entry, markup) {
        var host = markup && markup.host ? " via " + markup.host : "";
        return (entry.author || (markup && markup.owner_username) || "Nexus PLM") + host;
    }

    return {
        COMMENT: COMMENT, OPEN: OPEN, RESOLVED: RESOLVED,
        APPLIED_PREFIX: APPLIED_PREFIX,
        toEntries: toEntries,
        planApply: planApply,
        appliedIdOf: appliedIdOf
    };
}));
