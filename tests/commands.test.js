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

test("signed out, what is live is what Word leaves live", () => {
    // The informational commands, and the five that are how you GET a document - pressing one
    // of those signs you in first, as Word's handlers do. Nothing item-scoped: with no session
    // there is no state, and a command that needs the item would only be refused.
    const live = commands.COMMANDS.filter((c) => commands.isEnabled(c, OUT)).map((c) => c.id);
    assert.deepStrictEqual(live.sort(), [
        "about", "connection", "help", "navigator", "new", "open", "search", "settings",
        "sign_in", "worklist"
    ]);
});

test("a command offered signed out still needs a session to run, and says so", () => {
    for (const id of ["new", "open", "search", "worklist", "settings"]) {
        assert.ok(commands.sessionOptional(commands.byId(id)), id);
        assert.ok(commands.needsSession(commands.byId(id)), id);
    }
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

test("Markup and Apply markups are offered for a document PLM knows, like every item command", () => {
    // ONLYOFFICE hands a plugin its comments (GetAllComments / AddComment), so both halves of
    // Word's Markup split button are real here.
    for (const id of ["markup", "apply_markups"]) {
        assert.ok(commands.isEnabled(commands.byId(id), MINE), id);
        assert.ok(commands.isEnabled(commands.byId(id), CHECKED_IN), id);
        assert.ok(!commands.isEnabled(commands.byId(id), IN), `${id} needs the document`);
    }
});

test("New Workflow and Release want the document checked in, as Word refuses them otherwise", () => {
    for (const id of ["new_workflow", "release"]) {
        assert.ok(commands.isEnabled(commands.byId(id), CHECKED_IN), id);
        assert.ok(!commands.isEnabled(commands.byId(id), MINE), id);
        assert.match(commands.disabledBecause(commands.byId(id), MINE), /Check the document in/);
    }
});

test("Reload Document is refused while the document is checked out to me, as Word refuses it", () => {
    // A copy checked out to you is the newer one; reloading would throw your work away.
    assert.ok(!commands.isEnabled(commands.byId("reload_document"), MINE));
    assert.match(commands.disabledBecause(commands.byId("reload_document"), MINE), /your copy is the newer one/);
    assert.ok(commands.isEnabled(commands.byId("reload_document"), CHECKED_IN));
    assert.ok(commands.isEnabled(commands.byId("reload_document"), THEIRS), "someone else may have saved a newer version");
});

test("Edit Values needs the document checked out to me, as Word insists", () => {
    assert.ok(commands.isEnabled(commands.byId("edit_values"), MINE));
    assert.ok(!commands.isEnabled(commands.byId("edit_values"), CHECKED_IN));
    assert.ok(!commands.isEnabled(commands.byId("edit_values"), THEIRS));
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

test("a button is the large kind, as every button on Word's ribbon is", () => {
    for (const item of toolbar.tab("asc.{TEST}", IN).tabs[0].items) {
        assert.strictEqual(item.type, toolbar.BIG_BUTTON, item.id);
        assert.ok(item.lockInViewMode, `${item.id} should grey in a read-only document`);
    }
});

test("Word's groups are drawn as groups: a separator opens every one but the first", () => {
    const items = toolbar.tab("asc.{TEST}", IN).tabs[0].items;
    const openers = items.filter((i) => i.separator).map((i) => i.id);
    const firsts = commands.GROUPS.slice(1)
        .map((g) => toolbar.buttonId(commands.inGroup(g).filter((c) => !c.parent)[0].id));
    assert.deepStrictEqual(openers, firsts);
    assert.ok(!items[0].separator, "nothing before the first group");
});

test("Word's split buttons carry their menus: Markup, Refresh Values and Help", () => {
    const items = toolbar.tab("asc.{TEST}", MINE).tabs[0].items;
    const menus = {};
    for (const item of items) {
        if (item.items) {
            assert.ok(item.split, `${item.id} has a menu but is not a split button`);
            menus[item.id] = item.items.map((m) => m.id);
        }
    }
    assert.deepStrictEqual(menus, {
        [toolbar.buttonId("markup")]: [toolbar.buttonId("apply_markups")],
        [toolbar.buttonId("refresh_values")]: [toolbar.buttonId("reload_document")],
        [toolbar.buttonId("help")]: [toolbar.buttonId("about"), toolbar.buttonId("connection")]
    });
});

test("a menu entry is never enabled while its button is greyed, because the editor greys them together", () => {
    // Measured: a disabled ButtonCustom takes its menu with it. So commands.js may only nest an
    // entry whose rule is at least as strict as its parent's, and this holds it.
    for (const context of [OUT, IN, CHECKED_IN, MINE, THEIRS]) {
        for (const item of toolbar.tab("asc.{TEST}", context).tabs[0].items) {
            for (const entry of item.items || []) {
                if (item.disabled) { assert.ok(entry.disabled, `${entry.id} enabled under greyed ${item.id}`); }
            }
        }
    }
});

test("a menu entry's click arrives under an id that maps back to its command", () => {
    for (const id of toolbar.clickIds()) {
        assert.ok(toolbar.commandFor(id), id);
    }
    assert.strictEqual(toolbar.commandFor(toolbar.buttonId("about")).id, "about");
    assert.strictEqual(toolbar.clickIds().length, commands.COMMANDS.length);
});

test("every command appears on the tab, in group order", () => {
    const payload = toolbar.tab("asc.{TEST}", IN);
    // Every command is either a button or an entry in a button's menu - nothing is missing and
    // nothing is drawn twice.
    const drawn = [];
    for (const item of payload.tabs[0].items) {
        drawn.push(item.id);
        for (const entry of item.items || []) { drawn.push(entry.id); }
    }
    assert.deepStrictEqual(drawn.slice().sort(), toolbar.clickIds().slice().sort());
    assert.strictEqual(payload.tabs[0].items.length, commands.topLevel().length);

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
        assert.match(item.icons, /^icons\/%theme-type%\(light\|dark\)\/[a-z_]+%scale%\(default\)\.png$/);
    }
});

test("every command has icon files on disk, light and dark, at the five scales the editor asks for", () => {
    // A pattern the editor cannot resolve draws a button with no image and no error - the same
    // silent failure the LibreOffice build check exists to catch. %scale%(default) expands to
    // 100, 125, 150, 175 and 200 percent, and the editor names the files itself
    // (Common.UI.iconsStr2IconsObj): x.png, x@1.25x.png, x@1.5x.png, x@1.75x.png, x@2x.png.
    const fs = require("node:fs");
    const path = require("node:path");
    for (const command of commands.COMMANDS) {
        for (const theme of ["light", "dark"]) {
            for (const scale of ["", "@1.25x", "@1.5x", "@1.75x", "@2x"]) {
                const file = path.join(__dirname, "..", "plugin", "icons", theme,
                                       `${command.id}${scale}.png`);
                assert.ok(fs.existsSync(file), `missing ${theme}/${command.id}${scale}.png`);
            }
        }
    }
});

test("an icon is the editor's own big-button size at every scale", () => {
    // 28px at 100%, as the AI plugin shipped with Desktop Editors 9.4.0 draws its big buttons.
    const fs = require("node:fs");
    const path = require("node:path");
    const size = (file) => {
        const b = fs.readFileSync(file);
        return [b.readUInt32BE(16), b.readUInt32BE(20)];
    };
    const expected = { "": 28, "@1.25x": 35, "@1.5x": 42, "@1.75x": 49, "@2x": 56 };
    for (const [scale, px] of Object.entries(expected)) {
        const file = path.join(__dirname, "..", "plugin", "icons", "light", `check_out${scale}.png`);
        assert.deepStrictEqual(size(file), [px, px], `check_out${scale}.png`);
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

// ── what each command posts ──────────────────────────────────────────────────

test("the browsers are told what this host can open and where to root", () => {
    // Word posts these too. A command that posts the wrong body is answered, refused, and looks
    // exactly like a command that did nothing at all.
    for (const id of ["new", "open", "save_as_existing"]) {
        const body = commands.bodyFor(commands.byId(id), IN);
        assert.strictEqual(body.root_base_type, "DocumentsBase", id);
        assert.match(body.file_extensions, /\.docx/, id);
    }
    assert.strictEqual(commands.bodyFor(commands.byId("open"), IN).stage_assembly, false,
        "a document has no assembly to load");
});

test("every item command carries the item", () => {
    const context = { signedIn: true, user: "admin",
                      state: { status: "checked_in", item_id: "rev-1", file_path: "C:/x.docx" } };
    for (const id of ["check_out", "release", "revise", "properties", "edit_values",
                      "refresh_values", "reload_document", "new_workflow", "change_owner"]) {
        assert.strictEqual(commands.bodyFor(commands.byId(id), context).item_id, "rev-1", id);
    }
});

test("an upload carries the path the host read for the document, not one PLM answered with", () => {
    // /plm/state answers no path; the host's own reading of the open file is the only source.
    const context = { signedIn: true, user: "admin", state: { item_id: "rev-1", status: "checked_out", checked_out_by: "admin" },
                      doc: { title: "NX1.docx", path: "C:\Nexus\Staging\NX1.docx" } };
    for (const id of ["save", "check_in", "save_as_new", "save_as_existing", "change_owner", "markup"]) {
        assert.strictEqual(commands.bodyFor(commands.byId(id), context).file_path, "C:\Nexus\Staging\NX1.docx", id);
    }
});

test("About names the host, so the service reports the right one", () => {
    const body = commands.bodyFor(commands.byId("about"), IN);
    assert.strictEqual(body.host_name, "ONLYOFFICE");
    assert.ok(body.addin_version);
});

test("Connection Status posts nothing, because it is a report and not a command", () => {
    assert.strictEqual(commands.bodyFor(commands.byId("connection"), IN), null);
});

test("every command that calls the service has a body for it", () => {
    for (const c of commands.COMMANDS) {
        if (!c.endpoint || c.id === "connection") { continue; }
        assert.notStrictEqual(commands.bodyFor(c, IN), undefined, `${c.id} has no body`);
    }
});
