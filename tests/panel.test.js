/*
 * What the panel shows, tested with no ONLYOFFICE and no browser.
 *
 * Node's own test runner and assert — no framework, because a scaffold that needs an install
 * before its tests run is a scaffold people stop running.
 */

const test = require("node:test");
const assert = require("node:assert");

const panel = require("../plugin/lib/panel.js");
const editors = require("../plugin/lib/editors.js");

const CHECKED_IN = { status: "checked_in", part_number: "NX00000042", revision: "B", type_name: "n5Doc" };
const MINE = { status: "checked_out", checked_out_by: "admin", part_number: "NX00000042", revision: "B" };
const THEIRS = { status: "checked_out", checked_out_by: "jdoe", part_number: "NX00000042", revision: "B" };

test("a document PLM knows is named by its part number", () => {
    assert.strictEqual(panel.headline(CHECKED_IN, { hasDocument: true }), "NX00000042");
});

test("a document PLM has never seen says so as a fact, not a failure", () => {
    assert.strictEqual(panel.headline(null, { hasDocument: true }), panel.NOT_IN_PLM);
});

test("no document open is a different thing from not in PLM", () => {
    assert.strictEqual(panel.headline(null, { hasDocument: false }), panel.NO_DOCUMENT);
});

test("signed out is reported as signed out, whatever the document is", () => {
    // Otherwise every document in the vault reads as "not in PLM" the moment a session expires,
    // which sends the user looking for a problem with their file.
    assert.strictEqual(panel.headline(CHECKED_IN, { signedIn: false }), panel.NOT_SIGNED_IN);
});

test("a status of unknown is not in PLM", () => {
    // The service answers unknown both for a file it has never seen and for one it could not
    // resolve. The panel can only honestly speak to what it can show.
    assert.strictEqual(panel.isInPlm({ status: "unknown" }), false);
});

test("every row is present even when there is nothing to put in it", () => {
    const rows = panel.rowsFor(null);
    assert.strictEqual(rows.length, panel.ROWS.length);
    assert.ok(rows.every(([, value]) => value === panel.ABSENT));
});

test("a missing value shows the absent mark rather than an empty cell", () => {
    const rows = panel.rowsFor({ status: "checked_in", part_number: "NX00000042" });
    const description = rows.find(([label]) => label === "Description");
    assert.strictEqual(description[1], panel.ABSENT);
});

test("the lifecycle word is shown the way a person says it", () => {
    const rows = panel.rowsFor(CHECKED_IN);
    assert.strictEqual(rows.find(([label]) => label === "Status")[1], "Checked in");
});

test("a checked-in document offers Check Out and not Check In", () => {
    const allowed = panel.enabledButtons(CHECKED_IN, "admin");
    assert.ok(allowed.includes("check_out"));
    assert.ok(!allowed.includes("check_in"));
});

test("a document checked out to me offers Check In and Save", () => {
    const allowed = panel.enabledButtons(MINE, "admin");
    assert.ok(allowed.includes("check_in"));
    assert.ok(allowed.includes("save"));
    assert.ok(!allowed.includes("check_out"));
});

test("a document checked out to somebody else offers neither", () => {
    // The service would refuse, and an offer that is always refused is worse than no offer.
    const allowed = panel.enabledButtons(THEIRS, "admin");
    assert.ok(!allowed.includes("check_in"));
    assert.ok(!allowed.includes("check_out"));
    assert.ok(!allowed.includes("save"));
});

test("nothing is offered for a document PLM does not know", () => {
    assert.deepStrictEqual(panel.enabledButtons(null, "admin"), []);
});

test("not knowing who is signed in never makes a document mine", () => {
    assert.ok(!panel.enabledButtons(MINE, null).includes("check_in"));
});

test("the panel names the editor it is in, so two open at once are told apart", () => {
    assert.strictEqual(panel.title(editors.WORD), "Nexus PLM — Document");
    assert.strictEqual(panel.title(editors.CELL), "Nexus PLM — Spreadsheet");
    assert.strictEqual(panel.title(editors.SLIDE), "Nexus PLM — Presentation");
});

test("an editor ONLYOFFICE has not told us about still gets a panel", () => {
    assert.strictEqual(panel.title("pdf"), "Nexus PLM");
});
