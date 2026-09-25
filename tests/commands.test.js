/*
 * The Nexus PLM tab: that it matches Word, and that a command is offered exactly when it can work.
 *
 * The enable rules are the point. Offering a command that the service will then refuse is worse
 * than not offering it — the user cannot tell a refusal from a bug — and greying one that would
 * have worked is just as wrong.
 */

const test = require("node:test");
const assert = require("node:assert");

const commands = require("../plugin/lib/commands.js");
const toolbar = require("../plugin/lib/toolbar.js");

const OUT = { signedIn: false };
const IN = { signedIn: true, user: "admin" };
const CHECKED_IN = { signedIn: true, user: "admin",
                     state: { status: "checked_in", part_number: "NX1" } };
const MINE = { signedIn: true, user: "admin",
               state: { status: "checked_out", checked_out_by: "admin", part_number: "NX1" } };
const THEIRS = { signedIn: true, user: "admin",
                 state: { status: "checked_out", checked_out_by: "jdoe", part_number: "NX1" } };

// ── the set, against Word ────────────────────────────────────────────────────

test("the tab is named as Word names it", () => {
    assert.strictEqual(commands.TAB, "Nexus PLM");
});

test("the groups are Word's, in Word's order", () => {
    assert.deepStrictEqual(commands.GROUPS, [
        "Account", "Data Management", "Navigation and View", "Tasks",
        "Workflow", "Attribute Exchange", "Settings"
    ]);
});

test("every command sits in a group the tab actually has", () => {
    // A command in an unlisted group would simply never be drawn, with nothing logged.
    for (const c of commands.COMMANDS) {
        assert.ok(commands.GROUPS.includes(c.group), `${c.id} is in "${c.group}"`);
    }
});

test("every command id is unique", () => {
    const ids = commands.COMMANDS.map((c) => c.id);
    assert.strictEqual(new Set(ids).size, ids.length);
});

test("Word's whole set is here", () => {
    // Read from PLMRibbon.xml in Nexus.PLM.Office.WordAddin, not from memory.
    for (const id of ["sign_in", "sign_out", "new", "open", "save", "save_as_new",
                      "save_as_existing", "navigator", "search", "properties", "check_out",
                      "check_in", "release", "revise", "markup", "apply_markups", "change_owner",
                      "worklist", "new_workflow", "edit_values", "refresh_values",
                      "reload_document", "settings", "help", "about", "connection"]) {
        assert.ok(commands.byId(id), `${id} is missing`);
    }
});

test("every command either calls the service or is explicitly local", () => {
    // A command with neither is one nothing will ever happen for.
    const local = ["navigator", "help"];
    for (const c of commands.COMMANDS) {
        if (local.includes(c.id)) { assert.strictEqual(c.endpoint, null, c.id); }
        else { assert.ok(c.endpoint, `${c.id} has no endpoint`); }
    }
});

// ── when a command may be pressed ────────────────────────────────────────────

test("signed out, only the things that report or explain are live", () => {
    const live = commands.COMMANDS.filter((c) => commands.isEnabled(c, OUT)).map((c) => c.id);
    assert.deepStrictEqual(live.sort(), ["about", "connection", "help", "navigator", "sign_in"]);
});

test("signing in lights up everything that does not need the document", () => {
    const live = commands.COMMANDS.filter((c) => commands.isEnabled(c, IN)).map((c) => c.id);
    assert.ok(live.includes("new") && live.includes("open") && live.includes("search"));
    assert.ok(live.includes("worklist") && live.includes("settings") && live.includes("sign_out"));
    assert.ok(!live.includes("sign_in"), "already signed in");
    // Nothing that needs to know the item, because we do not know it yet.
    for (const c of commands.COMMANDS.filter(commands.needsDocument)) {
        assert.ok(!live.includes(c.id), `${c.id} needs a document`);
    }
});

test("a checked-in document offers Check Out and not Check In", () => {
    assert.ok(commands.isEnabled(commands.byId("check_out"), CHECKED_IN));
    assert.ok(!commands.isEnabled(commands.byId("check_in"), CHECKED_IN));
    assert.ok(!commands.isEnabled(commands.byId("save"), CHECKED_IN));
});

test("a document checked out to me offers Check In and Save", () => {
    assert.ok(commands.isEnabled(commands.byId("check_in"), MINE));
    assert.ok(commands.isEnabled(commands.byId("save"), MINE));
    assert.ok(!commands.isEnabled(commands.byId("check_out"), MINE));
});

test("a document checked out to somebody else offers neither", () => {
    // The service would refuse, and an offer that is always refused is worse than no offer.
    for (const id of ["check_in", "check_out", "save"]) {
        assert.ok(!commands.isEnabled(commands.byId(id), THEIRS), id);
    }
});

test("Markup is shown but not offered, because this host cannot do it yet", () => {
    // Word has it; ONLYOFFICE's comments are not read yet. Shown and explained beats missing.
    for (const id of ["markup", "apply_markups"]) {
        assert.ok(commands.byId(id), `${id} should still be on the tab`);
        assert.ok(!commands.isEnabled(commands.byId(id), MINE), `${id} cannot work yet`);
        assert.match(commands.disabledBecause(commands.byId(id), MINE), /does not read comments/);
    }
});

test("every greyed command says why", () => {
    // A disabled button with no reason is a dead end the user pokes at.
    for (const context of [OUT, IN, CHECKED_IN, MINE, THEIRS]) {
        for (const c of commands.COMMANDS) {
            const why = commands.disabledBecause(c, context);
            if (!commands.isEnabled(c, context)) {
                assert.ok(why && why.length > 10, `${c.id} greyed with no reason`);
            } else {
                assert.strictEqual(why, null, `${c.id} is enabled but gives a reason`);
            }
        }
    }
});

test("signed out, the reason is to sign in rather than something about the document", () => {
    assert.match(commands.disabledBecause(commands.byId("check_out"), OUT), /Sign in/);
});

// ── the payload handed to ONLYOFFICE ─────────────────────────────────────────

test("the payload carries the plugin's guid and one tab", () => {
    const payload = toolbar.tab("asc.{TEST}", IN);
    assert.strictEqual(payload.guid, "asc.{TEST}");
    assert.strictEqual(payload.tabs.length, 1);
    assert.strictEqual(payload.tabs[0].text, "Nexus PLM");
});

test("every command appears on the tab, in group order", () => {
    const payload = toolbar.tab("asc.{TEST}", IN);
    assert.strictEqual(payload.tabs[0].items.length, commands.COMMANDS.length);

    const order = payload.tabs[0].items.map((i) => toolbar.commandFor(i.id).group);
    const firstSeen = [...new Set(order)];
    assert.deepStrictEqual(firstSeen, commands.GROUPS);
});

test("a button id maps back to its command, and nothing else does", () => {
    assert.strictEqual(toolbar.commandFor(toolbar.buttonId("check_out")).id, "check_out");
    assert.strictEqual(toolbar.commandFor("someone_elses_button"), null);
    assert.strictEqual(toolbar.commandFor(undefined), null);
});

test("button ids are prefixed so they cannot collide with the editor's own", () => {
    const payload = toolbar.tab("asc.{TEST}", IN);
    for (const item of payload.tabs[0].items) {
        assert.ok(item.id.startsWith("nexusplm_"), item.id);
    }
});

test("a greyed button carries its reason as the tooltip", () => {
    const payload = toolbar.tab("asc.{TEST}", OUT);
    const checkOut = payload.tabs[0].items.find((i) => i.id === toolbar.buttonId("check_out"));
    assert.strictEqual(checkOut.disabled, true);
    assert.match(checkOut.hint, /Sign in/);
});

test("an enabled button's tooltip is just its label", () => {
    const payload = toolbar.tab("asc.{TEST}", IN);
    const open = payload.tabs[0].items.find((i) => i.id === toolbar.buttonId("open"));
    assert.strictEqual(open.disabled, false);
    assert.strictEqual(open.hint, "Open");
});

test("the Navigator is a toggle, because it shows and hides a panel", () => {
    const payload = toolbar.tab("asc.{TEST}", IN);
    const nav = payload.tabs[0].items.find((i) => i.id === toolbar.buttonId("navigator"));
    assert.strictEqual(nav.enableToggle, true);
});

test("button states cover every command", () => {
    assert.strictEqual(toolbar.buttonStates(IN).length, commands.COMMANDS.length);
});

// ── icons ────────────────────────────────────────────────────────────────────

test("every button carries an icon pattern", () => {
    const payload = toolbar.tab("asc.{TEST}", IN);
    for (const item of payload.tabs[0].items) {
        assert.ok(item.icons, `${item.id} has no icon`);
        assert.match(item.icons, /^icons\/%theme-type%/);
    }
});

test("every command has icon files on disk, light and dark, default and 2x", () => {
    // A pattern the editor cannot resolve draws a button with no image and no error — the same
    // silent failure the LibreOffice build check exists to catch.
    const fs = require("node:fs");
    const path = require("node:path");
    for (const command of commands.COMMANDS) {
        for (const theme of ["light", "dark"]) {
            for (const scale of ["", "@2x"]) {
                const file = path.join(__dirname, "..", "plugin", "icons", theme,
                                       `${command.id}${scale}.png`);
                assert.ok(fs.existsSync(file), `missing ${theme}/${command.id}${scale}.png`);
            }
        }
    }
});

test("an icon file is a real PNG, not an empty placeholder", () => {
    const fs = require("node:fs");
    const path = require("node:path");
    const file = path.join(__dirname, "..", "plugin", "icons", "light", "check_out.png");
    const bytes = fs.readFileSync(file);
    assert.ok(bytes.length > 100, "suspiciously small");
    assert.deepStrictEqual([...bytes.subarray(0, 8)], [137, 80, 78, 71, 13, 10, 26, 10]);
});
