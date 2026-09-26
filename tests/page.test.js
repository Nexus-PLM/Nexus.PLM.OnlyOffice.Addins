/*
 * The pages load their modules in an order that actually works, and the manifest says what the
 * editor needs to hear.
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
const CONFIG = JSON.parse(fs.readFileSync(path.join(PLUGIN, "config.json"), "utf8"));

function scriptsOf(page) {
    const html = fs.readFileSync(path.join(PLUGIN, page), "utf8");
    return [...html.matchAll(/src="((?:lib\/)?[\w.-]+\.js)"/g)].map((m) => m[1]);
}

/** Every page the config.json variations load. */
const PAGES = ["background.html", "index.html", "about.html"];

/** The two working pages: the resident half and the panel. Both run commands. */
const WORKING = ["background.html", "index.html"];

/** What each module reads at load time, from its own UMD factory call. */
function dependenciesOf(file) {
    const source = fs.readFileSync(path.join(PLUGIN, file), "utf8");
    const deps = [];
    for (const m of source.matchAll(/require\("\.\/([\w.-]+\.js)"\)/g)) {
        deps.push("lib/" + m[1]);
    }
    return deps;
}

/** What each glue file reads from window at load time — the globals the lib modules define. */
const GLOBALS = {
    "lib/editors.js": "NexusPlmEditors", "lib/client.js": "NexusPlmClient", "lib/commands.js": "NexusPlmCommands",
    "lib/paths.js": "NexusPlmPaths",
    "lib/panel.js": "NexusPlmPanel", "lib/toolbar.js": "NexusPlmToolbar", "lib/host.js": "NexusPlmHost",
    "lib/markup.js": "NexusPlmMarkup", "editor.js": "NexusPlmEditor", "runner.js": "NexusPlmRunner", "ui.js": "NexusPlmUi"
};

test("every page loads its modules in a workable order", () => {
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

test("the working pages load every module the plugin has", () => {
    // A module nobody loads is one whose global is missing at runtime, which shows up as an
    // empty panel or a dead tab rather than as an error.
    const onDisk = fs.readdirSync(path.join(PLUGIN, "lib")).filter((f) => f.endsWith(".js")).map((f) => "lib/" + f);
    for (const page of WORKING) {
        const order = scriptsOf(page);
        for (const file of onDisk) {
            assert.ok(order.includes(file), `${file} is never loaded by ${page}`);
        }
    }
});

test("a glue file comes after every global it reads at load time", () => {
    // editor.js, runner.js, ui.js and background.js read window.NexusPlm* the moment they run.
    for (const page of WORKING) {
        const order = scriptsOf(page);
        for (const file of order) {
            if (file.startsWith("lib/") || file === "secret.js") { continue; }
            const source = fs.readFileSync(path.join(PLUGIN, file), "utf8");
            for (const [dep, global] of Object.entries(GLOBALS)) {
                if (dep === file || !new RegExp("window\\." + global + "\\b").test(source)) { continue; }
                assert.ok(order.indexOf(dep) !== -1 && order.indexOf(dep) < order.indexOf(file),
                    `${page}: ${file} reads ${global} but ${dep} is not loaded before it`);
            }
        }
    }
});

test("the runner is shared: the tab's page and the panel's page both load it", () => {
    for (const page of WORKING) {
        assert.ok(scriptsOf(page).includes("runner.js"), `${page} does not load runner.js`);
    }
});

test("ui.js and plugin.js come last on the panel's page", () => {
    const order = scriptsOf("index.html");
    assert.strictEqual(order[order.length - 1], "plugin.js");
    assert.strictEqual(order[order.length - 2], "ui.js");
});

test("every page a variation names actually exists", () => {
    // A variation pointing at a missing page shows up as an entry in the dropdown that does
    // nothing at all when chosen.
    for (const variation of CONFIG.variations) {
        assert.ok(fs.existsSync(path.join(PLUGIN, variation.url)),
            `${variation.url} is named by a variation but is not there`);
        assert.ok(PAGES.includes(variation.url), `${variation.url} is not covered by these tests`);
    }
});

// ── the manifest ─────────────────────────────────────────────────────────────

test("the first variation is a resident background plugin, so the tab has a frame to reach", () => {
    // Measured on Desktop Editors 9.4.0: a variation with no type and isVisual false is parsed
    // as PluginType.Invisible, a one-shot action torn down as soon as init returns. Its tab
    // survived that with nothing behind it, and every button did nothing.
    const first = CONFIG.variations[0];
    assert.strictEqual(first.url, "background.html");
    assert.strictEqual(first.type, "background");
    assert.strictEqual(first.isVisual, false);
});

test("the background variation declares the toolbar click event it lives on", () => {
    // What the shipped AI plugin declares; the editor delivers a tab click as this event.
    assert.ok(CONFIG.variations[0].events.includes("onToolbarMenuClick"));
});

test("the Navigator panel docks beside the document, as Word's does", () => {
    const panel = CONFIG.variations.find((v) => v.url === "index.html");
    assert.ok(panel, "no panel variation");
    assert.strictEqual(panel.description, "Navigator");
    assert.strictEqual(panel.type, "panel");
    assert.strictEqual(panel.isVisual, true);
});

test("the dropdown offers a way in and a way out", () => {
    const descriptions = CONFIG.variations.map((v) => v.description).join(" | ");
    // ONLYOFFICE labels the FIRST variation "Start" and turns it into "Stop" once running - it
    // supplies that itself, so there must be no Start or Stop variation of our own. Checked
    // against LanguageTool, whose dropdown is exactly Start + About.
    assert.ok(!/^(Start|Stop)$/m.test(descriptions), "Start/Stop are the editor's to label, not ours");
    assert.strictEqual(CONFIG.variations[CONFIG.variations.length - 1].description, "About");
});

test("the plugin asks for a build that has AddToolbarMenuItem", () => {
    // minVersion is the one thing the editor checks before loading; the toolbar API arrived
    // with 8.2 (the AI plugin declares the same floor).
    assert.strictEqual(CONFIG.minVersion, "8.2.0");
});
