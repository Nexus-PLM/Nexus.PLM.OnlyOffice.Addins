/*
 * Comments to markups and back, tested on the shapes ONLYOFFICE actually uses.
 *
 * The comment shape is what the SDK's own serializer writes (sdk-all.js 9.4.0): {Id, Data} with
 * Data {Text, Time, UserName, QuoteText, Solved, UserData, Replies}. Time is milliseconds since
 * the epoch, as a string.
 */

const test = require("node:test");
const assert = require("node:assert");

const markup = require("../plugin/lib/markup.js");

const T = "1758844800000"; // 2025-09-26T00:00:00Z
const COMMENT = { Id: "c1", Data: { Text: "Wrong torque", Time: T, UserName: "admin", QuoteText: "45 Nm", Solved: false,
                                    Replies: [{ Text: "Agreed", Time: T, UserName: "jdoe" }] } };
const RESOLVED = { Id: "c2", Data: { Text: "Typo", Time: T, UserName: "admin", QuoteText: "", Solved: true, Replies: [] } };

// ── to PLM ───────────────────────────────────────────────────────────────────

test("a comment becomes one entry keyed by the editor's own comment id", () => {
    const entries = markup.toEntries([RESOLVED]);
    assert.strictEqual(entries.length, 1);
    const e = entries[0];
    assert.strictEqual(e.external_id, "c2");
    assert.strictEqual(e.parent_id, null);
    assert.strictEqual(e.kind, markup.COMMENT);
    assert.strictEqual(e.author, "admin");
    assert.strictEqual(e.text, "Typo");
    assert.strictEqual(e.status, markup.RESOLVED);
    assert.strictEqual(e.created_at, "2025-09-26T00:00:00.000Z");
});

test("a reply is an entry of its own, under its comment", () => {
    const entries = markup.toEntries([COMMENT]);
    assert.strictEqual(entries.length, 2);
    assert.strictEqual(entries[0].status, markup.OPEN);
    assert.strictEqual(entries[0].quote, "45 Nm");
    assert.strictEqual(entries[1].parent_id, "c1");
    assert.strictEqual(entries[1].external_id, "c1/0");
    assert.strictEqual(entries[1].author, "jdoe");
});

test("the entry keys are the service's snake_case, so they deserialise", () => {
    const keys = Object.keys(markup.toEntries([RESOLVED])[0]).sort();
    assert.deepStrictEqual(keys, ["author", "created_at", "external_id", "kind", "ordinal", "parent_id", "quote", "status", "text"]);
});

test("a comment that came from PLM is not sent back to it", () => {
    const applied = { Id: "c9", Data: { Text: "From FreeCAD", Time: T, UserName: "jdoe via FreeCAD", UserData: markup.APPLIED_PREFIX + "entry-7" } };
    assert.deepStrictEqual(markup.toEntries([applied, RESOLVED]).map((e) => e.external_id), ["c2"]);
});

test("a time the editor did not set still gives a valid timestamp", () => {
    const e = markup.toEntries([{ Id: "c3", Data: { Text: "x", UserName: "a" } }])[0];
    assert.ok(!isNaN(Date.parse(e.created_at)));
});

// ── from PLM ─────────────────────────────────────────────────────────────────

const PLM = [{
    markup_id: "m1", host: "Word", owner_username: "jdoe",
    entries: [
        { id: "e1", external_id: "w-1", parent_id: null, kind: "Comment", author: "jdoe", created_at: "2025-09-26T00:00:00Z",
          text: "Check the flange", quote: "flange", status: "Open" },
        { id: "e2", external_id: "w-1/0", parent_id: "e1", kind: "Comment", author: "admin", created_at: "2025-09-26T01:00:00Z",
          text: "Will do", status: "Open" },
        { id: "e3", external_id: "c2", parent_id: null, kind: "Comment", author: "admin", created_at: "2025-09-26T00:00:00Z",
          text: "Typo", status: "Resolved" }
    ]
}];

test("a reviewer's markup becomes a comment, replies included, named for where it came from", () => {
    const additions = markup.planApply(PLM, []);
    assert.strictEqual(additions.length, 2);
    const a = additions[0];
    assert.strictEqual(a.Text, "Check the flange");
    assert.strictEqual(a.UserName, "jdoe via Word");
    assert.strictEqual(a.QuoteText, "flange");
    assert.strictEqual(a.Solved, false);
    assert.strictEqual(a.Time, String(Date.parse("2025-09-26T00:00:00Z")));
    assert.strictEqual(a.UserData, markup.APPLIED_PREFIX + "e1");
    assert.strictEqual(a.Replies.length, 1);
    assert.strictEqual(a.Replies[0].Text, "Will do");
    assert.strictEqual(additions[1].Solved, true);
});

test("an entry this document produced is not applied back into it", () => {
    // c2 is one of this document's own comments: its external id IS the comment id.
    const additions = markup.planApply(PLM, [RESOLVED]);
    assert.deepStrictEqual(additions.map((a) => a.Text), ["Check the flange"]);
});

test("an entry a previous Apply already wrote is not written twice", () => {
    const already = { Id: "c9", Data: { Text: "Check the flange", UserData: markup.APPLIED_PREFIX + "e1" } };
    const additions = markup.planApply(PLM, [already]);
    assert.deepStrictEqual(additions.map((a) => a.Text), ["Typo"]);
});

test("nothing to apply is an empty plan, not an error", () => {
    assert.deepStrictEqual(markup.planApply([], []), []);
    assert.deepStrictEqual(markup.planApply(null, null), []);
});
