/*
 * The page loads its modules in an order that actually works.
 *
 * There is no bundler and no module loader in a plugin frame: every file is a plain <script>, and
 * one that reads another's global while it is being evaluated must come after it. Getting this
 * wrong does not fail loudly — the module throws, its global is never defined, and the panel comes
 * up COMPLETELY EMPTY with nothing on screen and nothing in any log to say why. That happened.
 */

const test = require("node:test");
const assert = require("node:assert");
const fs = require("node:fs");
const path = require("node:path");

const PLUGIN = path.join(__dirname, "..", "plugin");

function scriptsOf(page) {
    const html = fs.readFileSync(path.join(PLUGIN, page), "utf8");
    return [...html.matchAll(/src="((?:lib\/)?[\w.-]+\.js)"/g)].map((m) => m[1]);
}

/** Every page the config.json variations load. */
const PAGES = ["background.html", "index.html", "about.html"];

/** The panel's own order, which the dependency tests below are about. */
const ORDER = scriptsOf("index.html");

/** What each module reads at load time, from its own UMD factory call. */
function dependenciesOf(file) {
    const source = fs.readFileSync(path.join(PLUGIN, file), "utf8");
    const deps = [];
    for (const m of source.matchAll(/require\("\.\/([\w.-]+\.js)"\)/g)) {
        deps.push("lib/" + m[1]);
    }
    return deps;
}

test("every module is loaded after the ones it reads at load time", () => {
    for (const file of ORDER) {
        if (!file.startsWith("lib/")) { continue; }
        const at = ORDER.indexOf(file);
        for (const dep of dependenciesOf(file)) {
            const depAt = ORDER.indexOf(dep);
            assert.notStrictEqual(depAt, -1, `${file} needs ${dep}, which the page never loads`);
            assert.ok(depAt < at,
                `${file} is loaded at ${at} but reads ${dep} at ${depAt} — the panel will be empty`);
        }
    }
});

test("the page loads every module the plugin has", () => {
    // A module nobody loads is one whose global is missing at runtime, which shows up as an
    // empty panel rather than as an error.
    const onDisk = fs.readdirSync(path.join(__dirname, "..", "plugin", "lib"))
        .filter((f) => f.endsWith(".js")).map((f) => "lib/" + f);
    for (const file of onDisk) {
        assert.ok(ORDER.includes(file), `${file} is never loaded by index.html`);
    }
});

test("ui.js and plugin.js come last, after everything they use", () => {
    assert.ok(ORDER.indexOf("ui.js") > ORDER.filter((f) => f.startsWith("lib/"))
        .map((f) => ORDER.indexOf(f)).reduce((a, b) => Math.max(a, b), -1));
    assert.strictEqual(ORDER[ORDER.length - 1], "plugin.js");
});

test("every page loads its modules in a workable order", () => {
    // background.html has the same trap as index.html: it builds the tab from commands.js the
    // moment toolbar.js is evaluated.
    for (const page of PAGES) {
        const order = scriptsOf(page);
        for (const file of order) {
            if (!file.startsWith("lib/")) { continue; }
            for (const dep of dependenciesOf(file)) {
                assert.ok(order.indexOf(dep) !== -1 && order.indexOf(dep) < order.indexOf(file),
                    `${page}: ${file} is loaded before ${dep}`);
            }
        }
    }
});

test("every page a variation names actually exists", () => {
    // A variation pointing at a missing page shows up as an entry in the dropdown that does
    // nothing at all when chosen.
    const config = JSON.parse(fs.readFileSync(path.join(PLUGIN, "config.json"), "utf8"));
    for (const variation of config.variations) {
        assert.ok(fs.existsSync(path.join(PLUGIN, variation.url)),
            `${variation.url} is named by a variation but is not there`);
        assert.ok(PAGES.includes(variation.url), `${variation.url} is not covered by these tests`);
    }
});

test("the dropdown offers a way in and a way out", () => {
    // A background half that does not start on its own needs the user to be able to start it -
    // and having started it, to stop it.
    const config = JSON.parse(fs.readFileSync(path.join(PLUGIN, "config.json"), "utf8"));
    const descriptions = config.variations.map((v) => v.description).join(" | ");
    // ONLYOFFICE labels the FIRST variation "Start" and turns it into "Stop" once running - it
    // supplies that itself, so there must be no Start or Stop variation of our own. Checked
    // against LanguageTool, whose dropdown is exactly Start + About.
    assert.ok(!/^(Start|Stop)$/m.test(descriptions),
        "Start/Stop are the editor's to label, not ours");
    assert.match(descriptions, /About/);

    // The VISUAL panel must be first. A background variation is a one-shot action: it runs, is
    // torn down, and its heartbeat stops - measured. The tab's registration survives that, but
    // no frame is left to receive a click, which is why every button did nothing. The panel is
    // the only frame that stays alive, so it is the one that owns the tab.
    // The background half is first because it is the only thing measured to put the tab up.
    // The panel is what runs commands, because it is the only frame that stays alive.
    assert.strictEqual(config.variations[0].url, "background.html");
    assert.strictEqual(config.variations[0].isVisual, false);
    assert.strictEqual(config.variations[config.variations.length - 1].description, "About");
});
