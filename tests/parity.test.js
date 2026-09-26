/*
 * Parity with the Word ribbon, rule by rule.
 *
 * Word is the reference host: its ribbon's enable rules live in `RibbonEnablement.IsEnabled`
 * (Nexus.PLM.Office.Core), and its handlers refuse a few more cases with a toast. This file
 * holds ONLYOFFICE against that table so a difference is either intended and named here, or a
 * failing test.
 *
 * The table below is Word's, transcribed from that switch. Where this host is deliberately
 * stricter the reason is the Word HANDLER that would have refused anyway — the user hears it
 * from a greyed button's tooltip instead of from a toast after a wasted click.
 */

const test = require("node:test");
const assert = require("node:assert");

const commands = require("../plugin/lib/commands.js");

/** What Word's ribbon answers, given the four facts it reduces a document to. */
function word(id, ctx) {
    switch (id) {
        case "sign_in": case "sign_out": case "about": case "connection": case "help":
        case "new": case "open": case "search": case "worklist": case "navigator": case "settings":
            return true;
        case "save_as_new":      return ctx.hasDocument && !ctx.registered;
        case "save_as_existing": return ctx.hasDocument;
        case "save": case "edit_values":   return ctx.registered && ctx.mine;
        case "reload_document":            return ctx.registered && !ctx.mine;
        default:                           return ctx.registered;
    }
}

/** The same four facts, as this host's command table takes them. */
function ours(ctx, signedIn) {
    var state = ctx.registered
        ? { status: ctx.mine ? "checked_out" : (ctx.theirs ? "checked_out" : "checked_in"),
            checked_out_by: ctx.mine ? "admin" : (ctx.theirs ? "jdoe" : null), part_number: "NX1" }
        : null;
    return { signedIn: signedIn !== false, user: signedIn === false ? null : "admin", state: state };
}

const CASES = {
    "not registered":        { hasDocument: true, registered: false, mine: false, theirs: false },
    "checked in":            { hasDocument: true, registered: true,  mine: false, theirs: false },
    "checked out to me":     { hasDocument: true, registered: true,  mine: true,  theirs: false },
    "checked out to jdoe":   { hasDocument: true, registered: true,  mine: false, theirs: true  }
};

/**
 * Where this host is stricter than Word's RIBBON, and the Word handler that refuses the same
 * case anyway — so the outcome matches and only the moment the user is told differs.
 */
const STRICTER = {
    // OnRelease: "Check the document in before releasing it."
    release:        ["checked out to me", "checked out to jdoe"],
    // OnNewWorkflow: "Check in the document before adding it to a workflow."
    new_workflow:   ["checked out to me", "checked out to jdoe"],
    // OnCheckInOut: one button whose label follows the lock; the other half always refuses.
    check_out:      ["checked out to me", "checked out to jdoe"],
    check_in:       ["checked in", "checked out to jdoe"],
    // CheckInAsync returns silently unless the lock is the user's own.
    save:           ["checked out to jdoe"],
    edit_values:    ["checked out to jdoe"],
    // Word's ribbon leaves Save As Existing on with no document registered or not; ours wants a
    // session, which every one of its callers needs anyway.
    save_as_existing: []
};

/**
 * Commands Word models as ONE control and this host as two, so their rules cannot line up
 * one-for-one. Word's Sign In button is always enabled because it is also the Sign Out button;
 * here they are separate, each live in its own half of the session. The last test in this file
 * pins that down.
 */
const TWO_CONTROLS_HERE = ["sign_in", "sign_out"];

test("every command's rule matches Word's, or is a named stricter case", () => {
    for (const [name, ctx] of Object.entries(CASES)) {
        const context = ours(ctx);
        for (const command of commands.COMMANDS) {
            if (TWO_CONTROLS_HERE.includes(command.id)) { continue; }
            const mine = commands.isEnabled(command, context);
            const theirs = word(command.id, ctx);
            if (mine === theirs) { continue; }
            const allowed = STRICTER[command.id] || [];
            assert.ok(!mine && allowed.includes(name),
                `${command.id} on a document ${name}: this host says ${mine}, Word says ${theirs}` +
                (mine ? " — MORE permissive than Word, which is a functional break" : " — stricter, and not a named case"));
        }
    }
});

test("Save As is refused once the document is already a PLM item, as Word refuses it", () => {
    // The break this test exists for: offering Save As on a registered document creates a
    // SECOND PLM item for one file. Word greys the button and its handler refuses again.
    const registered = ours(CASES["checked out to me"]);
    assert.ok(!commands.isEnabled(commands.byId("save_as_new"), registered));
    assert.match(commands.disabledBecause(commands.byId("save_as_new"), registered), /already registered/);

    const fresh = ours(CASES["not registered"]);
    assert.ok(commands.isEnabled(commands.byId("save_as_new"), fresh), "a document PLM has never seen may be registered");
});

test("signed out, a command that needs a session climbs the ladder rather than refusing", () => {
    // Word: every handler calls EnsureSignedInAsync first, so New pressed while signed out
    // signs in and then runs. `runner.js` does the same for anything needsSession() marks.
    for (const id of ["new", "open", "search", "worklist", "settings", "save", "check_in", "properties"]) {
        assert.ok(commands.needsSession(commands.byId(id)), id);
    }
    // These must never trigger it: they are what works when nothing else does, and Sign In IS
    // the ladder.
    for (const id of ["sign_in", "help", "about", "connection", "navigator"]) {
        assert.ok(!commands.needsSession(commands.byId(id)), id);
    }
});

test("signed out, the same commands are live as in Word", () => {
    // Word's ribbon has no session dimension at all: New, Open, Search, My Worklist, Current
    // Settings and the informational commands are live signed out, and the handler signs you in.
    // Everything item-scoped is greyed because the document's state is unknown until there is a
    // session. Greying the first five would cost the user a second click Word does not ask for.
    const out = ours(CASES["not registered"], false);
    const live = commands.COMMANDS.filter((c) => commands.isEnabled(c, out)).map((c) => c.id).sort();
    assert.deepStrictEqual(live, [
        "about", "connection", "help", "navigator", "new", "open", "search", "settings",
        "sign_in", "worklist"
    ]);
});

test("every command that opens a type or vault browser declares what this host can open", () => {
    // Word sends `file_extensions` on New, Save As, Save As Existing and Open, and `host_name`
    // on the two that show the New Object form. Without them the form offers types whose
    // template this host cannot open, and its chip reads "Has template" rather than naming the
    // host — measured: the Save As form listed 25 types where New listed 24.
    const ctx = { signedIn: true, user: "admin", state: null, doc: { title: "d.docx", path: "C:\\p\\d.docx" } };
    for (const id of ["new", "save_as_new", "save_as_existing", "open"]) {
        const body = commands.bodyFor(commands.byId(id), ctx);
        assert.ok(body.file_extensions, `${id} does not declare file_extensions`);
        assert.match(body.file_extensions, /\.docx/);
    }
    for (const id of ["new", "save_as_new"]) {
        assert.strictEqual(commands.bodyFor(commands.byId(id), ctx).host_name, "ONLYOFFICE", id);
    }
});

test("the differences from Word that remain are the ones we chose", () => {
    // A list, so adding a difference means editing this test on purpose.
    // 1. Check Out and Check In are separate buttons; Word has one split button whose label
    //    follows the lock. The editor greys a split button's menu with the button, and Release
    //    must stay reachable while Check In is not.
    assert.ok(commands.byId("check_out") && commands.byId("check_in"));
    // 2. Sign In and Sign Out are separate buttons; Word has one toggle.
    assert.strictEqual(commands.byId("sign_in").when, commands.SIGNED_OUT);
    assert.strictEqual(commands.byId("sign_out").when, commands.SIGNED_IN);
    // 3. Word has a brand button that opens About; here About is in the Help menu.
    assert.strictEqual(commands.byId("about").parent, "help");
});
