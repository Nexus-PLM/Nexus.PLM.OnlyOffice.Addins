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

/** The plugin's own storage, which every frame of it shares. */
function fakeStorage() {
    const held = {};
    return {
        getItem: (k) => (Object.prototype.hasOwnProperty.call(held, k) ? held[k] : null),
        setItem: (k, v) => { held[k] = String(v); },
        removeItem: (k) => { delete held[k]; }
    };
}

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
        localStorage: o.localStorage || fakeStorage(),
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

test("the notice is a modal window of the plugin's own, and its message is left in storage", () => {
    const storage = fakeStorage();
    const w = recordingWindow();
    const f = frame({ PluginWindow: w.PluginWindow, localStorage: storage });

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

    assert.strictEqual(s.url, "notice.html");
    assert.deepStrictEqual(JSON.parse(storage.getItem("nexusplm.notice.v1")),
        { message: "Nexus PLM is not running. Start the tray.", severity: "error" });
});

test("the notice window's url carries NO fragment", () => {
    // MEASURED, Desktop Editors 9.4.0: a plugin window whose url carries a `#` never loads its
    // page. The window opens, shows the editor's own "Loading" spinner for ever, logs nothing,
    // and cannot be closed - so the window written to break a silence became worse than the
    // silence. The Navigator's own `index.html`, with no fragment, has always loaded. This is
    // the whole reason the message goes through storage, so it is held on its own.
    const w = recordingWindow();
    const f = frame({ PluginWindow: w.PluginWindow });

    f.NexusPlmEditor.showNotice("anything at all", "error");

    assert.ok(!w.shown[0].settings.url.includes("#"),
        "a plugin window url with a fragment never loads its page");
});

test("the message is in storage BEFORE the window is opened", () => {
    // The page reads it as it loads, so writing it afterwards would be a race the user sees as
    // the fallback line.
    const storage = fakeStorage();
    let atOpen = null;
    class Watching {
        constructor() { this.id = 1; }
        show() { atOpen = storage.getItem("nexusplm.notice.v1"); }
        close() {}
    }
    const f = frame({ PluginWindow: Watching, localStorage: storage });

    f.NexusPlmEditor.showNotice("Nexus PLM is not running.", "error");

    assert.ok(atOpen, "nothing was in storage when the window opened");
    assert.strictEqual(JSON.parse(atOpen).message, "Nexus PLM is not running.");
});

test("storage that refuses does not stop the window opening", () => {
    // A window saying the fallback line beats no window at all.
    const refusing = {
        getItem() { throw new Error("denied"); },
        setItem() { throw new Error("denied"); }
    };
    const w = recordingWindow();
    const f = frame({ PluginWindow: w.PluginWindow, localStorage: refusing });

    assert.strictEqual(f.NexusPlmEditor.showNotice("anything", "error"), true);
    assert.strictEqual(w.shown.length, 1);
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

/** The notice page, loaded against a given storage; answers the box it drew into. */
function noticePage(storage) {
    const box = { textContent: "", className: "np-notice" };
    const context = {
        JSON, String,
        localStorage: storage,
        document: { getElementById: (id) => (id === "np-notice" ? box : null) },
        Asc: {}
    };
    context.window = context;
    vm.createContext(context);
    vm.runInContext(fs.readFileSync(path.join(PLUGIN, "notice.js"), "utf8"), context, { filename: "notice.js" });
    return { box, context };
}

test("the notice page reads what showNotice left in storage", () => {
    // The two halves of one contract, exercised together: what the tab writes, the page reads.
    const storage = fakeStorage();
    const w = recordingWindow();
    frame({ PluginWindow: w.PluginWindow, localStorage: storage })
        .NexusPlmEditor.showNotice("Nexus PLM is not running.", "warning");

    const { box, context } = noticePage(storage);

    assert.strictEqual(box.textContent, "Nexus PLM is not running.");
    assert.strictEqual(box.className, "np-notice np-notice-warning");
    assert.strictEqual(typeof context.Asc.plugin.init, "function", "the SDK expects init on every page");
});

test("the notice page never comes up blank", () => {
    // An empty window is the failure this page exists to avoid, so with nothing to read it says
    // the one thing it still knows.
    const { box } = noticePage(fakeStorage());

    assert.match(box.textContent, /tray application/);
    assert.ok(box.textContent.length > 20, "a blank notice is worse than no notice");
});

test("the notice page survives storage it cannot read", () => {
    const { box } = noticePage({ getItem() { throw new Error("denied"); } });
    assert.match(box.textContent, /tray application/);
});
