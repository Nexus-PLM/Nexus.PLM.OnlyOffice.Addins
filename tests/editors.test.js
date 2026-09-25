/*
 * The three editors, and the promise that one plugin serves all of them.
 *
 * These tests exist to catch the failure this repo is most likely to have: a fourth copy of the
 * plugin appearing, or config.json and the code disagreeing about which editors are supported.
 */

const test = require("node:test");
const assert = require("node:assert");
const fs = require("node:fs");
const path = require("node:path");

const editors = require("../plugin/lib/editors.js");

const CONFIG = JSON.parse(
    fs.readFileSync(path.join(__dirname, "..", "plugin", "config.json"), "utf8")
);

test("all three editors are described", () => {
    assert.deepStrictEqual(
        editors.supported().map((e) => e.key),
        ["word", "cell", "slide"]
    );
});

test("config.json declares exactly the editors the code knows", () => {
    // If these drift, the plugin either loads in an editor it cannot describe or refuses to load
    // in one it can. Neither fails loudly inside ONLYOFFICE.
    const declared = CONFIG.variations[0].EditorsSupport.slice().sort();
    const known = editors.supported().map((e) => e.key).sort();
    assert.deepStrictEqual(declared, known);
});

test("there is ONE variation serving all three, not one variation each", () => {
    // The whole architecture of this repo. Three variations would be three plugins in a trench
    // coat, and they would drift the way the Excel and PowerPoint add-ins did.
    assert.strictEqual(CONFIG.variations.length, 1);
});

test("the guid is in the form ONLYOFFICE requires", () => {
    // A malformed guid makes the plugin fail to register with nothing logged.
    assert.match(CONFIG.guid, /^asc\.\{[0-9A-Fa-f-]{36}\}$/);
});

test("each editor drives the OOXML format its host actually writes", () => {
    const byKey = Object.fromEntries(editors.supported().map((e) => [e.key, e]));
    assert.strictEqual(byKey.word.extension, ".docx");
    assert.strictEqual(byKey.cell.extension, ".xlsx");
    assert.strictEqual(byKey.slide.extension, ".pptx");
});

test("a template extension is never the document extension", () => {
    // A template staged under its own extension made LibreOffice open "Untitled 1"; the staged
    // copy has to carry the document extension. Same formats here, so the same trap.
    for (const e of editors.supported()) {
        assert.notStrictEqual(e.templateExtension, e.extension);
    }
});

test("an unknown editor is described as nothing rather than guessed at", () => {
    assert.strictEqual(editors.describe("pdf"), null);
    assert.strictEqual(editors.describe(null), null);
    assert.strictEqual(editors.describe(undefined), null);
});

test("the editor type is matched however ONLYOFFICE cases it", () => {
    assert.strictEqual(editors.describe("WORD").key, "word");
});
