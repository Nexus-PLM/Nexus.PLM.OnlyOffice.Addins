/*
 * The installer, held against the plugin it installs.
 *
 * An installer is the one part of this repository that cannot be exercised by running it: it
 * writes into ONLYOFFICE's own folders on a real machine. So what CAN be checked without running
 * it is checked here, and it is not a formality — every rule below is one where being wrong
 * produces no error at all, just a plugin that never appears or never works.
 *
 * The folder name above all. ONLYOFFICE matches a plugin folder to its config by the guid, and a
 * folder named anything else is either ignored or registered twice. Nothing logs that.
 */

const test = require("node:test");
const assert = require("node:assert");
const fs = require("node:fs");
const path = require("node:path");

const ROOT = path.join(__dirname, "..");
const ISS = fs.readFileSync(
    path.join(ROOT, "installer", "Nexus.PLM.OnlyOffice.Addin.iss"), "utf8");
const CONFIG = JSON.parse(fs.readFileSync(path.join(ROOT, "plugin", "config.json"), "utf8"));

/** The value of an ISPP `#define NAME "value"` line. */
function define(name) {
    const match = ISS.match(new RegExp('^#define\\s+' + name + '\\s+"([^"]*)"', "m"));
    return match ? match[1] : null;
}

test("the plugin is installed under the folder name ONLYOFFICE will look for", () => {
    // The guid without its "asc." prefix, and with the leading brace doubled because a single one
    // opens an Inno constant. Get this wrong and the editor silently ignores the plugin.
    const expected = "{" + CONFIG.guid.replace(/^asc\./, "");
    assert.strictEqual(define("PluginFolder"), expected);
});

test("it installs into the per-user plugin folder, not the one under Program Files", () => {
    // The Program Files copy needs elevation and is replaced by the next ONLYOFFICE update. The
    // user folder also has its own `v1`, so the `../v1/plugins.js` every plugin page loads still
    // resolves from there.
    assert.match(define("PluginsRoot"),
        /^\{localappdata\}\\ONLYOFFICE\\DesktopEditors\\data\\sdkjs-plugins$/);
    assert.match(ISS, /^PrivilegesRequired=lowest$/m);
});

test("the installer's version is the plugin's version", () => {
    // Two files, one number. They drift the moment nothing compares them.
    assert.strictEqual(define("AppVersion"), CONFIG.version);
});

test("every file the plugin ships is carried, and secret.js is not", () => {
    // A recursive copy of plugin\, so a file added to the plugin needs no installer change - which
    // is the point, because an installer listing files one by one is an installer that forgets one.
    assert.match(ISS, /^Source: "\.\.\\plugin\\\*";/m);
    assert.match(ISS, /Flags: recursesubdirs createallsubdirs/);

    // secret.js is per machine and written at install time. One shipped in the payload would hand
    // every seat the same proof of locality, which is the whole thing the secret exists to prevent.
    assert.match(ISS, /Excludes: "secret\.js"/);
});

test("the install writes secret.js, because without it every command is refused", () => {
    // The plugin draws its tab either way. The failure is 403 on every call, which reads as a
    // broken add-in rather than a missing step, so the installer also says so when it cannot.
    assert.match(ISS, /window\.NexusPlmSecret/);
    assert.match(define("SecretSource"), /browser-addin-secret\.txt$/);
    assert.ok(/if not SecretWasWritten then\s+MsgBox\(/.test(ISS),
        "a missing secret must be reported, not left to look like a broken plugin");
});

test("uninstalling removes this plugin's folder and nothing around it", () => {
    // Its parent holds every other ONLYOFFICE plugin, the ones shipped with the editor included.
    const uninstall = ISS.match(/^Type: filesandordirs; Name: "([^"]+)"/m);
    assert.ok(uninstall, "the plugin folder is removed on uninstall");
    assert.match(uninstall[1], /PluginTarget|sdkjs-plugins/);
    assert.ok(!/Name: "\{#PluginsRoot\}"/.test(ISS),
        "the plugins root itself must never be deleted");
});

test("the editor is closed for the install, because it reads its plugin list once at startup", () => {
    // Installing underneath a running editor leaves the user looking at a window that will never
    // show the tab, and concluding the installer failed.
    assert.match(ISS, /^CloseApplications=yes$/m);
});

test("the plugin page loads the secret the installer writes", () => {
    // The two halves of the same mechanism, in different files and different languages. If the
    // page stops loading it, the installer is writing a file nobody reads.
    const page = fs.readFileSync(path.join(ROOT, "plugin", "index.html"), "utf8");
    const background = fs.readFileSync(path.join(ROOT, "plugin", "background.html"), "utf8");
    assert.match(page, /secret\.js/);
    assert.match(background, /secret\.js/);
});
