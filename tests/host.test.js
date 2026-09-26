/*
 * The host's half of every command, as data: what ONLYOFFICE has to do around a service call.
 *
 * Taken from Word's ribbon handlers, which are the most complete host, and held here so that a
 * command's follow-through cannot depend on which button it was pressed from — and so that the
 * rules Word learned the hard way (a cancelled dialog is not news; a new object's lock is what
 * the Engine says it is, not what was hoped) are not relearned.
 */

const test = require("node:test");
const assert = require("node:assert");

const commands = require("../plugin/lib/commands.js");
const host = require("../plugin/lib/host.js");

const ME = { user: "admin" };
const by = (id) => commands.byId(id);
const types = (effects) => effects.map((e) => e.type);

// ── before the call: the file on disk ────────────────────────────────────────

test("a command that uploads the file is refused while the document has unsaved changes", () => {
    const why = host.refuse(by("check_in"), { path: "C:\\x\\NX1.docx", saved: true, changes: 3 });
    assert.match(why, /unsaved changes/);
    assert.match(why, /Check In/);
});

test("a document never saved to a file cannot be uploaded", () => {
    assert.match(host.refuse(by("save"), { path: "C:\\x\\NX1.docx", saved: false, changes: 0 }), /Save the document to a file/);
});

test("a document whose path is not known cannot be uploaded, and the reason says how to fix that", () => {
    // The editor tells a plugin the file's name and not its path (measured; lib/paths.js), and
    // an upload hands the service a path.
    const why = host.refuse(by("check_in"), { title: "NX1.docx", path: null, saved: null, changes: null });
    assert.match(why, /Open it through Nexus PLM/);
    assert.match(why, /Check In/);
});

test("a saved, clean document is not refused", () => {
    assert.strictEqual(host.refuse(by("save"), { path: "C:\\x\\NX1.docx", saved: true, changes: 0 }), null);
});

test("when the editor cannot say whether the file is saved, the command goes ahead and the service decides", () => {
    // Refusing on a guess is worse than a refusal the service can explain.
    assert.strictEqual(host.refuse(by("check_in"), { path: "C:\\x\\NX1.docx", saved: null, changes: null }), null);
});

test("a command that does not upload the file never asks about it", () => {
    for (const id of ["new", "open", "check_out", "release", "properties", "worklist"]) {
        assert.strictEqual(host.refuse(by(id), { path: null, saved: false, changes: 9 }), null, id);
    }
});

test("exactly the commands that hand the service a file path are marked as uploading", () => {
    const uploading = commands.COMMANDS.filter((c) => c.saves).map((c) => c.id).sort();
    assert.deepStrictEqual(uploading, ["change_owner", "check_in", "save", "save_as_existing", "save_as_new"]);
});

// ── after the answer ─────────────────────────────────────────────────────────

test("a cancelled dialog is the user's own decision and produces nothing", () => {
    assert.deepStrictEqual(host.effectsFor(by("new"), { success: false, cancelled: true }, ME), []);
    assert.deepStrictEqual(host.effectsFor(by("open"), { success: false, error: "User cancelled" }, ME), []);
});

test("a service that cannot be reached is reported, in its own words", () => {
    const effects = host.effectsFor(by("save"), { success: false, unreachable: true, error: "Nexus PLM is not running." }, ME);
    assert.deepStrictEqual(types(effects), [host.SAY]);
    assert.strictEqual(effects[0].severity, "error");
    assert.match(effects[0].message, /not running/);
});

test("a refusal the service already announced is not announced twice", () => {
    // The service raises its own toast for a refusal; Word posting one as well showed two.
    assert.deepStrictEqual(host.effectsFor(by("check_out"), { success: false, error: "Checked out by jdoe" }, ME), []);
});

test("a body the service could not bind is reported as this plugin's own mistake, naming the field", () => {
    const effects = host.effectsFor(by("check_in"), { httpStatus: 400, title: "One or more validation errors occurred.", errors: { FilePath: ["required"] } }, ME);
    assert.strictEqual(effects.length, 1);
    assert.match(effects[0].message, /rejected the request \(FilePath\)/);
});

test("silence with no error at all is reported as no response", () => {
    const effects = host.effectsFor(by("release"), { success: false }, ME);
    assert.strictEqual(effects.length, 1);
    assert.match(effects[0].message, /No response/);
});

test("New opens the staged file and remembers nothing here: the new tab is its own document", () => {
    const effects = host.effectsFor(by("new"),
        { success: true, plm_object_id: "obj-1", part_number: "NX1", file_path: "C:\\stage\\NX1.docx", checked_out: true }, ME);
    assert.deepStrictEqual(types(effects), [host.OPEN]);
    assert.strictEqual(effects[0].path, "C:\\stage\\NX1.docx");
});

test("New with no staged file says so instead of opening nothing", () => {
    const effects = host.effectsFor(by("new"), { success: true, plm_object_id: "obj-1", part_number: "NX1" }, ME);
    assert.deepStrictEqual(types(effects), [host.SAY]);
    assert.match(effects[0].message, /NX1/);
});

test("Open and Search open the file the service staged", () => {
    for (const [id, answer] of [
        ["open", { success: true, file_path: "C:\\stage\\NX1.docx", item_id: "rev-1" }],
        ["search", { success: true, file_path: "C:\\stage\\NX1.docx", plm_object_id: "obj-1" }]
    ]) {
        const effects = host.effectsFor(by(id), answer, ME);
        assert.deepStrictEqual(types(effects), [host.OPEN], id);
    }
});

test("Search that picked an item with no file just names it", () => {
    const effects = host.effectsFor(by("search"), { success: true, plm_object_id: "obj-1", part_number: "NX1", type_name: "n5Doc" }, ME);
    assert.deepStrictEqual(types(effects), [host.SAY]);
    assert.match(effects[0].message, /NX1 \(n5Doc\)/);
});

test("Save As binds this document to the item it just became, with the lock the Engine granted", () => {
    const granted = host.effectsFor(by("save_as_new"),
        { success: true, plm_object_id: "obj-2", checked_out: true, attribute_mappings: { Title: "Spec" } }, ME);
    assert.deepStrictEqual(types(granted), [host.BIND, host.VALUES]);
    assert.strictEqual(granted[0].itemId, "obj-2");
    assert.strictEqual(granted[0].status, "checked_out");
    assert.strictEqual(granted[0].owner, "admin");
    assert.deepStrictEqual(granted[1].values, { Title: "Spec" });

    // Word used to stamp "checked out" unconditionally; the answer is the truth.
    const refused = host.effectsFor(by("save_as_new"), { success: true, plm_object_id: "obj-2", checked_out: false }, ME);
    assert.strictEqual(refused[0].status, "checked_in");
    assert.strictEqual(refused[0].owner, null);
});

test("Check In releases the lock and carries the new revision", () => {
    const effects = host.effectsFor(by("check_in"), { success: true, item_id: "rev-2", revision: "B" }, ME);
    assert.deepStrictEqual(effects, [{ type: host.BIND, itemId: "rev-2", status: "checked_in", owner: null, revision: "B" }]);
});

test("Revise with a staged file opens the new revision beside this one and says so", () => {
    const effects = host.effectsFor(by("revise"),
        { success: true, item_id: "rev-3", revision: "C", file_path: "C:\\stage\\NX1.docx", checked_out: true }, ME);
    assert.deepStrictEqual(types(effects), [host.OPEN, host.SAY]);
    assert.match(effects[1].message, /Revision C/);
});

test("Revise with no file makes this document the new revision, checked out to me", () => {
    const effects = host.effectsFor(by("revise"),
        { success: true, item_id: "rev-3", revision: "C", checked_out: true, attribute_mappings: { Rev: "C" } }, ME);
    assert.deepStrictEqual(types(effects), [host.BIND, host.VALUES]);
    assert.strictEqual(effects[0].status, "checked_out");
    assert.strictEqual(effects[0].owner, "admin");
});

test("Edit Values writes the values back only when something was saved", () => {
    assert.deepStrictEqual(host.effectsFor(by("edit_values"), { success: true, saved: false, attribute_mappings: { A: "1" } }, ME), []);
    const effects = host.effectsFor(by("edit_values"), { success: true, saved: true, attribute_mappings: { A: "1" } }, ME);
    assert.deepStrictEqual(effects, [{ type: host.VALUES, values: { A: "1" } }]);
});

test("Refresh Values writes what PLM sent and changes nothing else", () => {
    const effects = host.effectsFor(by("refresh_values"), { success: true, attribute_mappings: { A: "1" }, revision: "B" }, ME);
    assert.deepStrictEqual(types(effects), [host.VALUES]);
});

test("Reload Document opens the current version beside this one, because a plugin cannot close its own document", () => {
    const effects = host.effectsFor(by("reload_document"), { success: true, file_path: "C:\\stage\\NX1.docx", revision: "B" }, ME);
    assert.deepStrictEqual(types(effects), [host.OPEN, host.SAY]);
    assert.match(effects[1].message, /new tab/);
});

test("Settings says when a restart is needed, and is otherwise silent", () => {
    assert.deepStrictEqual(host.effectsFor(by("settings"), { saved: true, success: true }, ME), []);
    const effects = host.effectsFor(by("settings"), { saved: true, success: true, restart_required: true }, ME);
    assert.match(effects[0].message, /restarted/);
});

test("the commands whose dialogs the service owns end with nothing for the host to do", () => {
    for (const id of ["check_out", "change_owner", "worklist", "new_workflow", "properties", "about", "sign_in", "sign_out", "markup"]) {
        assert.deepStrictEqual(host.effectsFor(by(id), { success: true }, ME), [], id);
    }
});

test("every command that can move the lock, the revision or the lifecycle is re-read afterwards", () => {
    for (const id of ["check_out", "check_in", "release", "revise", "change_owner", "sign_in", "sign_out", "save_as_new", "save_as_existing"]) {
        assert.ok(host.movesState(by(id)), id);
    }
    for (const id of ["properties", "worklist", "settings", "about", "help", "refresh_values"]) {
        assert.ok(!host.movesState(by(id)), id);
    }
});
