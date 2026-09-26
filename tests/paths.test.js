/*
 * Where a document is on disk, remembered from what the service handed over.
 *
 * The editor tells a plugin the file's NAME and not its path (measured; see lib/paths.js), so
 * every path the service answers with is remembered by name, in storage every tab shares.
 */

const test = require("node:test");
const assert = require("node:assert");

const paths = require("../plugin/lib/paths.js");

function memoryStore() {
    const m = new Map();
    return { getItem: (k) => (m.has(k) ? m.get(k) : null), setItem: (k, v) => m.set(k, String(v)), removeItem: (k) => m.delete(k) };
}

test("a file the service staged is found again by its name", () => {
    const store = memoryStore();
    paths.remember(store, "C:\\Nexus\\Staging\\NX00000042.docx");
    assert.strictEqual(paths.lookup(store, "NX00000042.docx"), "C:\\Nexus\\Staging\\NX00000042.docx");
});

test("the name is matched without case, because Windows does", () => {
    const store = memoryStore();
    paths.remember(store, "C:\\Nexus\\Staging\\NX00000042.docx");
    assert.strictEqual(paths.lookup(store, "nx00000042.DOCX"), "C:\\Nexus\\Staging\\NX00000042.docx");
});

test("a name never opened here is looked for in the service's staging folder once that is known", () => {
    const store = memoryStore();
    assert.strictEqual(paths.lookup(store, "EM-00000038-DOC.docx"), null, "nothing known yet");
    paths.remember(store, "C:\\Nexus\\Staging\\NX00000042.docx");
    assert.strictEqual(paths.lookup(store, "EM-00000038-DOC.docx"), "C:\\Nexus\\Staging\\EM-00000038-DOC.docx");
});

test("the exact file wins over the staging folder", () => {
    const store = memoryStore();
    paths.remember(store, "C:\\Nexus\\Staging\\NX00000042.docx");
    paths.remember(store, "D:\\elsewhere\\EM-00000038-DOC.docx");
    assert.strictEqual(paths.lookup(store, "EM-00000038-DOC.docx"), "D:\\elsewhere\\EM-00000038-DOC.docx");
});

test("forward slashes are kept as forward slashes", () => {
    const store = memoryStore();
    paths.remember(store, "/home/marc/staging/NX1.docx");
    assert.strictEqual(paths.lookup(store, "NX2.docx"), "/home/marc/staging/NX2.docx");
});

test("what is remembered survives a fresh read, as it must across editor tabs", () => {
    const store = memoryStore();
    paths.remember(store, "C:\\Nexus\\Staging\\NX00000042.docx");
    const raw = store.getItem(paths.KEY);
    const again = memoryStore();
    again.setItem(paths.KEY, raw);
    assert.strictEqual(paths.lookup(again, "NX00000042.docx"), "C:\\Nexus\\Staging\\NX00000042.docx");
});

test("a broken store is an empty one, not an exception", () => {
    const store = memoryStore();
    store.setItem(paths.KEY, "{not json");
    assert.strictEqual(paths.lookup(store, "NX1.docx"), null);
    paths.remember(store, "C:\\Nexus\\Staging\\NX1.docx");
    assert.strictEqual(paths.lookup(store, "NX1.docx"), "C:\\Nexus\\Staging\\NX1.docx");
    assert.strictEqual(paths.lookup(null, "NX1.docx"), null);
});

test("an empty name remembers nothing and finds nothing", () => {
    const store = memoryStore();
    paths.remember(store, "");
    paths.remember(store, "C:\\Nexus\\Staging\\");
    assert.strictEqual(paths.lookup(store, ""), null);
    assert.strictEqual(store.getItem(paths.KEY), null);
});

test("clear forgets everything", () => {
    const store = memoryStore();
    paths.remember(store, "C:\\Nexus\\Staging\\NX1.docx");
    paths.clear(store);
    assert.strictEqual(paths.lookup(store, "NX1.docx"), null);
});
