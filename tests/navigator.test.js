/*
 * The docked Navigator: the tree it builds, and what may be done to the row you pick.
 *
 * The same pane Word docks beside the document, so these tests are written against the same
 * rules its React version holds (`src/embed/folderTreeModel.ts`, `lockOwner.ts`, and the action
 * buttons in `NavigatorPane.tsx`). Where a rule differs it is because this host can do something
 * Word's WebView2 cannot, or the reverse, and the test says which.
 */

const test = require("node:test");
const assert = require("node:assert");

const nav = require("../plugin/lib/navigator.js");

/** A folder as the service sends it. */
function wireFolder(id, name, parent, order, count) {
    return { folder_id: id, name: name, parent_id: parent, sort_order: order || 0,
             object_count: count || 0, has_badge: (count || 0) > 0, is_system: false };
}

const FOLDERS = [
    wireFolder("root", "Top", null, 1, 0),
    wireFolder("b", "Bravo", "root", 2, 3),
    wireFolder("a", "Alpha", "root", 1, 0),
    wireFolder("a1", "Alpha One", "a", 1, 5),
    wireFolder("mine", "My Working", null, 0, 5)
].map(nav.folderFromWire);

// ── the wire ─────────────────────────────────────────────────────────────────

test("a folder is read out of the service's snake_case, with a null parent at the root", () => {
    const f = nav.folderFromWire({ folder_id: "f1", name: "Drawings", parent_id: null,
                                   sort_order: 3, object_count: 12, has_badge: true, is_system: true });
    assert.deepStrictEqual(f, { folderId: "f1", name: "Drawings", parentId: null, isSystem: true,
                                sortOrder: 3, objectCount: 12, hasBadge: true });
});

test("an item is read the same way, and an absent lock is null rather than missing", () => {
    const i = nav.itemFromWire({ plm_object_id: "o1", revision_id: "r1", part_number: "NX1",
                                 type_name: "n5Doc", revision: "B", status: "In-work",
                                 modified_at: "2026-09-26T00:00:00Z" });
    assert.strictEqual(i.plmObjectId, "o1");
    assert.strictEqual(i.partNumber, "NX1");
    assert.strictEqual(i.revision, "B");
    assert.strictEqual(i.checkedOutBy, null);
    assert.strictEqual(i.description, null);
});

test("a wire row with nothing in it does not throw, and reads as empty", () => {
    assert.strictEqual(nav.folderFromWire(null).folderId, "");
    assert.strictEqual(nav.itemFromWire(undefined).partNumber, "");
});

// ── the tree ─────────────────────────────────────────────────────────────────

test("the tree nests by parent and sorts siblings by order, then name", () => {
    const tree = nav.buildTree(FOLDERS);
    assert.deepStrictEqual(tree.map((n) => n.folder.name), ["My Working", "Top"]);
    const top = tree[1];
    assert.deepStrictEqual(top.children.map((n) => n.folder.name), ["Alpha", "Bravo"]);
    assert.deepStrictEqual(top.children[0].children.map((n) => n.folder.name), ["Alpha One"]);
});

test("a folder whose parent is missing is promoted to the root, never dropped", () => {
    // Dropping it is how a tree silently loses a branch, and the pane may be a user's only view
    // of the vault. Word's buildTree keeps the same rule.
    const orphan = nav.folderFromWire(wireFolder("x", "Orphan", "gone", 9, 1));
    const tree = nav.buildTree(FOLDERS.concat([orphan]));
    assert.ok(tree.some((n) => n.folder.name === "Orphan"), "the orphan is at the root");
    assert.strictEqual(tree.length, 3);
});

test("a folder that claims itself as its parent does not build an endless tree", () => {
    const loop = nav.folderFromWire(wireFolder("self", "Loop", "self", 1, 0));
    const tree = nav.buildTree([loop]);
    assert.deepStrictEqual(tree.map((n) => n.folder.name), ["Loop"]);
});

test("ancestorsOf names every folder to expand, and stops on a cycle", () => {
    assert.deepStrictEqual(nav.ancestorsOf(FOLDERS, "a1"), ["a", "root"]);
    assert.deepStrictEqual(nav.ancestorsOf(FOLDERS, "root"), []);

    const cyclic = [nav.folderFromWire(wireFolder("p", "P", "q", 1, 0)),
                    nav.folderFromWire(wireFolder("q", "Q", "p", 1, 0))];
    assert.deepStrictEqual(nav.ancestorsOf(cyclic, "p"), ["q"]);
});

test("only the expanded branches are drawn, each at its own depth", () => {
    const tree = nav.buildTree(FOLDERS);
    const collapsed = nav.flatten(tree, {}, 0);
    assert.deepStrictEqual(collapsed.map((r) => r.folder.name), ["My Working", "Top"]);

    const open = nav.flatten(tree, { root: true, a: true }, 0);
    assert.deepStrictEqual(open.map((r) => [r.folder.name, r.depth]), [
        ["My Working", 0], ["Top", 0], ["Alpha", 1], ["Alpha One", 2], ["Bravo", 1]
    ]);
    assert.strictEqual(open.find((r) => r.folder.name === "Alpha").hasChildren, true);
    assert.strictEqual(open.find((r) => r.folder.name === "Bravo").hasChildren, false);
});

// ── locks ────────────────────────────────────────────────────────────────────

test("a lock is mine whatever the case or padding, and never when either side is unknown", () => {
    assert.ok(nav.isMyLock("Admin", "admin"));
    assert.ok(nav.isMyLock("  admin ", "admin"));
    assert.ok(!nav.isMyLock("jdoe", "admin"));
    assert.ok(!nav.isMyLock(null, "admin"));
    assert.ok(!nav.isMyLock("admin", null));
});

// ── the four actions ─────────────────────────────────────────────────────────

const FREE = nav.itemFromWire({ plm_object_id: "o1", part_number: "NX1", revision: "A", status: "In-work" });
const MINE = nav.itemFromWire({ plm_object_id: "o2", part_number: "NX2", checked_out_by: "admin" });
const THEIRS = nav.itemFromWire({ plm_object_id: "o3", part_number: "NX3", checked_out_by: "jdoe" });

test("with nothing picked, nothing is offered", () => {
    assert.deepStrictEqual(nav.actionsFor(null, "admin", null),
        { open: false, checkout: false, checkin: false, properties: false });
    assert.match(nav.whyDisabled("open", null, "admin", null), /Pick a document/);
});

test("a document nobody holds offers Open, Check Out and Properties", () => {
    const a = nav.actionsFor(FREE, "admin", null);
    assert.deepStrictEqual(a, { open: true, checkout: true, checkin: false, properties: true });
});

test("a document somebody else holds cannot be checked out, and says who has it", () => {
    const a = nav.actionsFor(THEIRS, "admin", null);
    assert.strictEqual(a.checkout, false);
    assert.strictEqual(a.checkin, false);
    assert.match(nav.whyDisabled("checkout", THEIRS, "admin", null), /Checked out to jdoe/);
});

test("Check In needs my own lock AND the document open here, as Word's host requires", () => {
    // The service uploads the working copy, so only the document this frame is docked to has a
    // path. Word refuses the same case in the same words - "is checked out to you but is not open
    // here. Open it first, then check it in." (`PLMRibbon.Navigator.cs`,
    // `NavigatorCheckInAsync`) - but it refuses AFTER the press, as a warning toast: its pane
    // enables the button on the lock alone. Greying it with that sentence as the tooltip says the
    // same thing before the click is spent, which is why this host does it that way.
    assert.strictEqual(nav.actionsFor(MINE, "admin", null).checkin, false);
    assert.match(nav.whyDisabled("checkin", MINE, "admin", null), /not open here/);
    assert.match(nav.whyDisabled("checkin", MINE, "admin", null), /Open it first/);

    assert.strictEqual(nav.actionsFor(MINE, "admin", "NX2").checkin, true, "open here: offered");
    assert.strictEqual(nav.actionsFor(MINE, "admin", "nx2").checkin, true, "matched without case");
});

test("Check In is refused for a lock that is not mine even when that document is open here", () => {
    assert.strictEqual(nav.actionsFor(THEIRS, "admin", "NX3").checkin, false);
    assert.match(nav.whyDisabled("checkin", THEIRS, "admin", "NX3"), /Checked out to jdoe/);
});

test("my own lock offers Check In rather than Check Out", () => {
    const a = nav.actionsFor(MINE, "admin", "NX2");
    assert.strictEqual(a.checkout, false);
    assert.match(nav.whyDisabled("checkout", MINE, "admin", "NX2"), /already have this checked out/);
});

// ── the band heading and the search box ──────────────────────────────────────

test("the band says what it is showing, and counts in the singular when there is one", () => {
    assert.strictEqual(nav.bandHeading({ items: [], folderId: null }), "Select a folder");
    assert.strictEqual(nav.bandHeading({ items: [], folderId: "f" }), "0 documents");
    assert.strictEqual(nav.bandHeading({ items: [FREE], folderId: "f" }), "1 document");
    assert.strictEqual(nav.bandHeading({ items: [FREE, MINE], folderId: "f" }), "2 documents");
    assert.strictEqual(nav.bandHeading({ items: [FREE], searching: true }), "1 match");
    assert.strictEqual(nav.bandHeading({ items: [FREE, MINE], searching: true }), "2 matches");
});

test("a query shorter than two characters is not worth a round trip", () => {
    assert.ok(!nav.isSearchable(""));
    assert.ok(!nav.isSearchable(" a "));
    assert.ok(nav.isSearchable("NX"));
    assert.strictEqual(nav.MIN_QUERY_LENGTH, 2, "Word's number");
});
