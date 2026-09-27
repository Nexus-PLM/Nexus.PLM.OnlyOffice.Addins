/*
 * editor.js against a stand-in editor: what Connection Status can see in each of the three
 * editors, and the one message the plugin can show without the tray.
 *
 * editor.js is a plugin-frame global, not a UMD module, so it is loaded into a vm context that
 * plays the frame: `window` is the context itself, so the `Api` and `Asc` a probe reaches for
 * inside callCommand are the same objects the test put there.
 */

const test = require("node:test");
const assert = require("node:assert");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const PLUGIN = path.join(__dirname, "..", "plugin");

/**
 * A frame with editor.js loaded. `api` is what the probe finds as `Api`; `plugin` overrides the
 * parts of Asc.plugin a test cares about.
 */
function frame(options) {
    const o = options || {};
    const context = {
        console,
        JSON, String, Array, Object, encodeURIComponent, decodeURIComponent,
        setTimeout, clearTimeout,
        localStorage: null,
        NexusPlmPaths: require("../plugin/lib/paths.js"),
        NexusPlmEditors: require("../plugin/lib/editors.js"),
        Api: o.api || {},
        Asc: {
            scope: {},
            plugin: Object.assign({
                info: { documentTitle: o.title || "NX1.docx" },
                // The real callCommand serialises fn into the editor's frame. Here the frame IS
                // this context, so calling it directly gives fn the same Api and Asc.scope.
                callCommand(fn, _a, _b, callback) { callback(fn()); },
                executeMethod(name, _args, callback) {
                    if (callback) { callback(name === "GetAllAddinFields" ? [] : null); }
                }
            }, o.plugin || {}),
            PluginWindow: o.PluginWindow
        }
    };
    context.window = context;
    context.self = context;
    vm.createContext(context);
    vm.runInContext(fs.readFileSync(path.join(PLUGIN, "editor.js"), "utf8"), context, { filename: "editor.js" });
    return context;
}

function probe(context, editorType) {
    return new Promise((resolve) => context.NexusPlmEditor.probe(editorType, resolve));
}

// ── what Connection Status can see, per editor ────────────────────────────────

test("in a Spreadsheet, Connection Status counts defined names and never asks for a Document", async () => {
    // Before: every editor was asked Api.GetDocument().GetAllContentControls(), which a
    // spreadsheet does not have, so the line said "no document builder" about a builder that
    // was there and answering. A spreadsheet's fields are its defined names.
    let askedForDocument = false;
    const api = {
        GetDocument() { askedForDocument = true; throw new Error("not a document"); },
        GetDefNames() { return [{}, {}, {}]; }
    };
    const line = await probe(frame({ api, title: "NX1.xlsx" }), "cell");

    assert.ok(!askedForDocument, "a spreadsheet was asked a document's question");
    assert.match(line, /^Spreadsheet: NX1\.xlsx/);
    assert.match(line, /3 defined name\(s\)/);
    assert.doesNotMatch(line, /no document builder/);
});

test("in a Presentation, Connection Status counts named shapes across every slide", async () => {
    const shape = (name) => ({ GetName: () => name });
    const api = {
        GetDocument() { throw new Error("not a document"); },
        GetPresentation() {
            return {
                GetSlidesCount: () => 2,
                GetSlideByIndex: (i) => ({
                    GetAllShapes: () => (i === 0 ? [shape("Title"), shape("")] : [shape("PartNumber")])
                })
            };
        }
    };
    const line = await probe(frame({ api, title: "NX1.pptx" }), "slide");

    assert.match(line, /^Presentation: NX1\.pptx/);
    assert.match(line, /2 named shape\(s\) on 2 slide\(s\)/);
});

test("in a Document, Connection Status still counts named content controls", async () => {
    // The Document line is what it was: this change added the other two, it did not move Word's.
    const control = (tag) => ({ GetAlias: () => "", GetTag: () => tag });
    const api = {
        GetDocument() {
            return {
                GetAllContentControls: () => [control("Title"), control(""), control("PartNumber")],
                GetCustomProperties: () => ({})
            };
        }
    };
    const line = await probe(frame({ api }), "word");

    assert.match(line, /^Document: NX1\.docx/);
    assert.match(line, /2 named content control\(s\), custom properties: yes/);
});

test("a builder that cannot answer is reported as the builder of THAT editor", async () => {
    // "no document builder" in a spreadsheet sent the reader to the wrong place.
    const api = { GetDefNames() { throw new Error("nope"); } };
    const line = await probe(frame({ api, title: "NX1.xlsx" }), "cell");
    assert.match(line, /no spreadsheet builder/);
});

test("an editor the plugin does not know still gets a line, without a probe", async () => {
    // lib/editors.js: an unknown editor is not an error. The probe must not throw either.
    let called = false;
    const api = { GetDocument() { called = true; return {}; } };
    const line = await probe(frame({ api, title: "NX1.vsdx" }), "visio");

    assert.ok(!called, "an unknown editor must not be asked a Document's questions");
    assert.match(line, /not readable in an editor this plugin does not know/);
});

// ── the one message that needs no tray ────────────────────────────────────────

/** A stand-in for Asc.PluginWindow that records what it was asked to show. */
function recordingWindow() {
    const shown = [];
    let nextId = 100;
    class PluginWindow {
        constructor() { this.id = nextId++; this.closed = false; }
        show(settings) { shown.push({ id: this.id, settings }); }
        close() { this.closed = true; }
    }
    return { PluginWindow, shown };
}

test("the notice is a modal window of the plugin's own, carrying the message in the fragment", () => {
    const w = recordingWindow();
    const f = frame({ PluginWindow: w.PluginWindow });

    const opened = f.NexusPlmEditor.showNotice("Nexus PLM is not running. Start the tray.", "error");

    assert.strictEqual(opened, true);
    assert.strictEqual(w.shown.length, 1);
    const s = w.shown[0].settings;
    assert.ok(s.isModal, "a notice the user has to see is modal");
    assert.ok(s.isVisual);
    // Array.from: an array made inside the frame has the frame's Array.prototype, and
    // deepStrictEqual holds that against it.
    assert.deepStrictEqual(Array.from(s.buttons, (b) => b.text), ["OK"]);
    assert.deepStrictEqual(Array.from(s.EditorsSupport), ["word", "cell", "slide"], "the notice serves every editor");

    // The URL fragment is the one part of a URL the editor's scheme handler never sees.
    assert.match(s.url, /^notice\.html#/);
    const payload = JSON.parse(decodeURIComponent(s.url.replace(/^notice\.html#/, "")));
    assert.deepStrictEqual(payload, { message: "Nexus PLM is not running. Start the tray.", severity: "error" });
});

test("the notice's OK button closes it, and a close box does the same", () => {
    const w = recordingWindow();
    const f = frame({ PluginWindow: w.PluginWindow });
    f.NexusPlmEditor.showNotice("hello", "info");
    const id = w.shown[0].id;

    // ONLYOFFICE reports a window's buttons to the OPENER's Asc.plugin.button(id, windowId).
    assert.strictEqual(f.NexusPlmEditor.windowButton(0, id), true);
    assert.strictEqual(f.NexusPlmEditor.windowButton(0, id), false, "already closed");

    f.NexusPlmEditor.showNotice("again", "info");
    const second = w.shown[1].id;
    f.NexusPlmEditor.windowClosed(second);
    assert.strictEqual(f.NexusPlmEditor.windowButton(0, second), false, "closed by its box");
});

test("a second notice replaces the first rather than stacking", () => {
    const w = recordingWindow();
    const f = frame({ PluginWindow: w.PluginWindow });
    f.NexusPlmEditor.showNotice("first", "info");
    f.NexusPlmEditor.showNotice("second", "info");
    assert.strictEqual(w.shown.length, 2);
});

test("a build with no plugin windows answers null, and nothing is thrown", () => {
    const f = frame({});
    assert.strictEqual(f.NexusPlmEditor.showNotice("hello", "info"), null);
    assert.strictEqual(f.NexusPlmEditor.showNotice("", "info"), false, "nothing to say");
});

test("the notice page reads the same fragment the window was opened with", () => {
    // notice.js is the other half of the contract: what showNotice encodes, it must decode.
    const payload = encodeURIComponent(JSON.stringify({ message: "Nexus PLM is not running.", severity: "warning" }));
    const box = { textContent: "", className: "np-notice" };
    const context = {
        JSON, String, decodeURIComponent,
        location: { hash: "#" + payload },
        document: { getElementById: (id) => (id === "np-notice" ? box : null) },
        Asc: {}
    };
    context.window = context;
    vm.createContext(context);
    vm.runInContext(fs.readFileSync(path.join(PLUGIN, "notice.js"), "utf8"), context, { filename: "notice.js" });

    assert.strictEqual(box.textContent, "Nexus PLM is not running.");
    assert.strictEqual(box.className, "np-notice np-notice-warning");
    assert.strictEqual(typeof context.Asc.plugin.init, "function", "the SDK expects init on every page");
});
