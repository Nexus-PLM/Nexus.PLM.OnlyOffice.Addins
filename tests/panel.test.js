/*
 * What the Navigator says when it cannot show the vault.
 *
 * Four different failures, said four different ways. Reporting all of them as "sign in" was
 * wrong: only one of them is fixed by signing in, and a plugin nobody installed properly would
 * send the user looking at permissions in PLM instead.
 */

const test = require("node:test");
const assert = require("node:assert");

const panel = require("../plugin/lib/panel.js");

test("nothing wrong is nothing said", () => {
    assert.strictEqual(panel.trouble({ signedIn: true, hasSecret: true }), null);
    assert.strictEqual(panel.trouble({}), null);
});

test("the missing secret is reported before anything else", () => {
    // With no secret every call is refused, so any other message would be a guess about a
    // service that never answered.
    assert.strictEqual(
        panel.trouble({ hasSecret: false, unreachable: true, refused: true, signedIn: false }),
        panel.NOT_INSTALLED);
    assert.match(panel.NOT_INSTALLED, /install step/);
});

test("a service that is not answering is not reported as signed out", () => {
    assert.strictEqual(panel.trouble({ unreachable: true, signedIn: false }), panel.NOT_RUNNING);
    assert.match(panel.NOT_RUNNING, /tray application/);
});

test("a refusal is its own message, and names the fix", () => {
    assert.strictEqual(panel.trouble({ refused: true }), panel.REFUSED);
    assert.match(panel.REFUSED, /install step/);
});

test("signed out says where Sign In actually is, because it is not in the pane", () => {
    assert.strictEqual(panel.trouble({ signedIn: false }), panel.NOT_SIGNED_IN);
    assert.match(panel.NOT_SIGNED_IN, /Nexus PLM tab/);
});

test("a read's own error is shown when the connection itself is fine", () => {
    assert.strictEqual(
        panel.trouble({ signedIn: true, error: "The folder tree could not be read." }),
        "The folder tree could not be read.");
});

test("the pane names the editor it is in, so two open at once are told apart", () => {
    assert.strictEqual(panel.title("word"), "Nexus PLM — Document");
    assert.strictEqual(panel.title("cell"), "Nexus PLM — Spreadsheet");
    assert.strictEqual(panel.title("slide"), "Nexus PLM — Presentation");
});

test("an editor ONLYOFFICE has not told us about still gets a title", () => {
    assert.strictEqual(panel.title(null), "Nexus PLM");
    assert.strictEqual(panel.title("diagram"), "Nexus PLM");
});
