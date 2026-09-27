/*
 * runner.js with a stand-in service and editor: what the runner says, and through what, when the
 * tray is there and when it is not.
 *
 * runner.js is a plugin-frame global, so it is loaded into a vm context holding the real lib
 * modules and a recording editor. The rules it applies live in lib/ and are tested there; what
 * is held here is the one thing the runner decides on its own - which way a message goes out.
 */

const test = require("node:test");
const assert = require("node:assert");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const PLUGIN = path.join(__dirname, "..", "plugin");
const commands = require("../plugin/lib/commands.js");

/** A service whose every call answers `answer`, remembering what it was told. */
function service(answer) {
    const calls = [];
    const reply = () => Promise.resolve(Object.assign({}, answer));
    return {
        calls,
        baseUrl: "http://localhost:5100",
        notify(message, severity) { calls.push({ notify: message, severity }); return reply(); },
        health() { calls.push({ health: true }); return reply(); },
        me: reply, autoLogin: reply, state: reply, command: reply,
        connect: reply, heartbeat: reply, disconnect: reply
    };
}

/** The frame, with a runner made for `editorType` against `client`. */
function runner(editorType, client) {
    const editor = {
        notices: [], probed: [],
        showNotice(message, severity) { this.notices.push({ message, severity }); return true; },
        probe(type, callback) { this.probed.push(type); callback("Fields: 1 defined name(s)"); },
        document(callback) { callback({ title: "NX1.xlsx", path: null, saved: null, changes: null, shell: false }); },
        save(callback) { callback(true); }
    };
    const context = {
        console, JSON, String, Array, Object, Promise, setTimeout, clearTimeout,
        NexusPlmCommands: commands,
        NexusPlmHost: require("../plugin/lib/host.js"),
        NexusPlmMarkup: require("../plugin/lib/markup.js"),
        NexusPlmClient: require("../plugin/lib/client.js"),
        NexusPlmPaths: require("../plugin/lib/paths.js"),
        NexusPlmEditor: editor
    };
    context.window = context;
    vm.createContext(context);
    vm.runInContext(fs.readFileSync(path.join(PLUGIN, "runner.js"), "utf8"), context, { filename: "runner.js" });
    return { runner: context.NexusPlmRunner.create({ editorType, client }), editor };
}

/** Let every promise the runner chained settle. */
const settle = () => new Promise((resolve) => setTimeout(resolve, 0));

// ── which way a message goes out ─────────────────────────────────────────────

test("with the tray running, a message is a Nexus toast and nothing else", async () => {
    const client = service({ success: true });
    const r = runner("cell", client);

    r.runner.say("Saved to Nexus PLM.", "success");
    await settle();

    assert.deepStrictEqual(client.calls, [{ notify: "Saved to Nexus PLM.", severity: "success" }]);
    assert.deepStrictEqual(r.editor.notices, [], "the toast appeared; a second message would say it twice");
});

test("with the tray closed, the message the toast could not carry is shown by the plugin itself", async () => {
    // The whole finding: the toast is drawn by the tray, so "the tray is not running" sent as a
    // toast reached nobody. Excel and PowerPoint had exactly this; Word answers with a message
    // box, and this is the plugin's equivalent.
    const client = service({ success: false, unreachable: true, error: "Nexus PLM is not running." });
    const r = runner("cell", client);

    r.runner.say("Nexus PLM is not running.", "error");
    await settle();

    assert.deepStrictEqual(r.editor.notices, [{ message: "Nexus PLM is not running.", severity: "error" }]);
});

test("a refusal the service answered is NOT repeated by the plugin", async () => {
    // Refused (403) or failed is the service's to explain - it owns the dialog - and it was
    // there to draw the toast. Only unreachable falls back.
    const client = service({ success: false, refused: true, httpStatus: 403 });
    const r = runner("cell", client);

    r.runner.say("Refused.", "warning");
    await settle();

    assert.deepStrictEqual(r.editor.notices, []);
});

// ── Connection Status ────────────────────────────────────────────────────────

test("Connection Status with the tray closed says so on screen, through the plugin's own window", async () => {
    const client = service({ success: false, unreachable: true, error: "Nexus PLM is not running. Start the tray." });
    const r = runner("slide", client);

    r.runner.runId("connection");
    await settle(); await settle();

    assert.strictEqual(r.editor.notices.length, 1);
    assert.match(r.editor.notices[0].message, /not running/);
    assert.strictEqual(r.editor.notices[0].severity, "error");
});

test("Connection Status asks the probe about the editor it is actually in", async () => {
    // It used to ask with no editor at all, and the probe asked the Document's questions of a
    // Spreadsheet and a Presentation, which is why both reported "no document builder".
    const client = service({ success: true });
    const r = runner("slide", client);

    r.runner.runId("connection");
    await settle(); await settle();

    assert.deepStrictEqual(r.editor.probed, ["slide"]);
    const toast = client.calls.find((c) => c.notify);
    assert.ok(toast, "the report is a toast when the tray is there");
    assert.match(toast.notify, /Connected to Nexus PLM on http:\/\/localhost:5100/);
    assert.match(toast.notify, /1 defined name/);
    assert.strictEqual(toast.severity, "warning", "not signed in, so the report is a warning");
});
