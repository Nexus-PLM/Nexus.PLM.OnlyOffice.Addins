# CLAUDE.md — Nexus.PLM.OnlyOffice.Addins

PLM inside the ONLYOFFICE editors: Document, Spreadsheet and Presentation. Read `README.md` first —
it is the human documentation and the source of truth for the layout.

**This repository is public and MIT.** Nothing private belongs in it: no internal hosts,
credentials, customer names, machine names or paths outside this repository.

## The shape of the work

| | |
|---|---|
| **The plugin** | JavaScript, loaded by the editor. No bundler and no module loader — the plugin frame is a plain page, so every file is a `<script>` tag in dependency order and each module is UMD so the tests can `require` the same file. |
| **The connector** | **There isn't one, and that is deliberate.** ONLYOFFICE reads and writes OOXML, so `Nexus.PLM.Office.*.Templates` already handles `.docx`/`.xlsx`/`.pptx` field discovery and value sync through OpenXml, with no Office and no COM, and the Vault already runs them. A connector here would be a second implementation of one format. |

**One plugin, three editors.** Every variation in `config.json` declares
`EditorsSupport: ["word", "cell", "slide"]`, and nothing outside `lib/editors.js` may branch on
which editor it is in. Tests assert that no variation serves fewer than all three, and that
`config.json` and the code agree about which editors exist — neither disagreement fails loudly
inside ONLYOFFICE.

Note it is **not** a rule that there is exactly one variation: the plugins shipped with Desktop
Editors 9.4.0 routinely carry a second for their About window, and this one will want one too.
What must hold is that no variation is for a single editor.

## `onlyofficeScheme` is why this works at all (measured, Desktop Editors 9.4.0)

`config.json` declares **`"onlyofficeScheme": true`**. Do not remove it. With it the editor serves
the plugin under its own registered scheme and the page's origin is `onlyoffice://plugin`; without
it the page is `file://`, its origin is *opaque*, and every `fetch` and `XMLHttpRequest` to the
Addin Service fails.

**How that was found, because the method matters more than the fact:** of the plugins shipped with
Desktop Editors, exactly three make network calls — AI, AI agent and DeepL — and exactly those
three declare the flag. `sdk-all.js` rewrites their base URL to `onlyoffice://plugin/<path>`.
Reading what already works beat three rounds of reasoning about CORS.

It is also what keeps the service narrow: a scheme the editor registers cannot be claimed by a web
page, so the AddinService allows `onlyoffice://plugin` and does **not** allow an opaque origin.

### The documentation describes this flag wrongly — trust the code

`api.onlyoffice.com` says `onlyofficeScheme` "specifies whether the plugin is included in the
server or desktop builds branded as ONLYOFFICE". That is not what it does, and anyone who reads
only the docs will not understand why this plugin needs it. Confirmed in ONLYOFFICE's own source:

- `sdkjs/common/Local/common.js` — `if (pluginsData[i]["onlyofficeScheme"]) { baseUrl =
  "onlyoffice://plugin/" + baseUrl; }`
- `desktop-sdk/ChromiumBasedEditors/lib/src/cefview.cpp` — the same rewrite, **unconditional**:
  there is no branding check despite the doc's wording, so it works in any desktop build.

**Still a prediction, not a measurement:** Docs in a browser over https, where the page is `https:`
and `http://localhost:5100` is mixed content that no CORS or scheme change fixes.

## The tab, the file and the shell — measured, Desktop Editors 9.4.0, Sep 25 2026

Every "because ONLYOFFICE does X" below was observed, not reasoned. Where a first guess was wrong
the wrong guess is recorded too, so nobody repeats it.

**The tab lives in a `type: "background"` variation.** The SDK (`CPluginVariation` in
`sdk-all-min.js`) maps `type: "background"` to `PluginType.Background` — resident, listed under
Plugins → Background plugins, started with every document unless the user stops it. A variation
with **no `type` and `isVisual: false` is `PluginType.Invisible`**: a one-shot action torn down as
soon as `init` returns. Its tab survived that with nothing behind it, which is why the tab used to
draw and then do nothing. The shipped AI plugin is the model. Clicks arrive through
`attachToolbarMenuClickEvent`; `UpdateToolbarMenuItem` with the same payload changes text, hint,
enabled state and menus of the buttons already drawn — never their icons.

**What a tab item can be** (from web-apps' `Common.UI.LayoutManager.addCustomControls`):
`type: "big-button"`, `separator: true` (a new group), `split: true` + `items: [{id, text}]` (a
menu; each entry clicks under its own id), `enableToggle`, `disabled`, `lockInViewMode`. A
disabled button greys its menu with it, so nothing may be nested under a button whose rule is
looser than its own — a test holds that. Captions drop off the later buttons when the window is
narrow; that is the editor's overflow, not a bug of ours.

**Icons:** a PATTERN, `icons/%theme-type%(light|dark)/<id>%scale%(default).png`. `default`
means 100, 125, 150, 175 and 200 percent and the editor names the files itself (`x.png`,
`x@1.25x.png`, `x@1.5x.png`, `x@1.75x.png`, `x@2x.png`; `Common.UI.iconsStr2IconsObj`). Big
buttons are 28px at 100%. `tools/make-icons.py` writes exactly that set from the 80px originals.

**Which file is open:** `Asc.plugin.info.documentTitle` — undocumented, present, the file's name.
The service resolves a staged file by its name (the part number), so the name alone answers
`/plm/state`. The **path** comes from the desktop shell's tool interface, the one the shipped AI
agent plugin uses: `AscDesktopEditor.callToolFunction("recent_files_reader")` lists every
recently opened file with its full path, open ones first; `lib/paths.js` remembers what the
service staged as the fallback. What does NOT work from a plugin frame, each tried:
`LocalFileGetSourcePath` / `LocalFileGetSaved` / `LocalFileGetOpenChangesCount` (answer for the
frame that opened the file, not this one), the `onupdaterecents` callback (never arrives),
`LocalFileOpen(path)` and `execCommand("open:recent", ...)` (accepted, open nothing), and any
shell or editor object from a `callCommand` script (that sandbox has `Api` and nothing else).

**Opening a staged file:** `callToolFunction("file_opener", {"path": ...})` — a new editor tab,
with its own copy of this plugin, which finds its item by name. Nothing about the opened file is
remembered in the tab that asked for it.

**Saving before an upload:** the builder's `Api.Save()` from a `callCommand` script writes the
local file (the file's timestamp moves). Whether the document is dirty cannot be read, so it is
asked for before every upload, as Word saves before every upload.

**Values:** `Api.GetDocument().GetCustomProperties().Add(name, value)` writes a custom property
and `Get` reads it back — measured through Refresh Values on a document whose header shows the
properties through `DOCPROPERTY` fields. Those fields keep showing their cached text: ONLYOFFICE's
`UpdateAllFields` does not re-evaluate `DOCPROPERTY`. The file's properties are right, which is
what the Vault's template connector and Word read; only the on-screen text is stale. And only
what the type's Data Model → Templates section MAPS comes back at all: the n5EMICAR answer was
seven attributes, and `NXDocumentClassification` was not among them because its mapping is
document → PLM. A placeholder that stays is first a mapping question, then a rendering one.

**Driven end to end on the tab, against the running service:** Sign In, Connection Status, Open
(dialog → staged file → new tab), Check In (save → upload → dialog → toast → state → tab
re-drawn), Check Out, Edit Values (dialog), Refresh Values (7 values written), Navigator (docks
left, buttons follow the state). Not yet exercised: New, Search, Save As, Save As Existing,
Release, Revise, Reload Document, Change Ownership, Markup, Apply markups, Worklist, New Workflow,
Properties, Settings, About — their bodies and follow-through are tested, not driven.

## Rules carried over from the other add-ins

These each exist because something went wrong in Word, Excel, PowerPoint, FreeCAD or LibreOffice.
They are cheaper to inherit than to relearn.

- **The service needs no changes for a new host.** A host declares what it can do and it travels
  with the request — `file_extensions`, `stage_assembly` (`false`: a document has no assembly). If
  a host seems to need a service change, stop and ask what it should be declaring instead.
- **A handler that opens a staged file must record which item it is.** Every PLM command keys off
  the file's path. Excel and PowerPoint skipped this in Search and every command stayed disabled on
  a file PLM had just handed over.
- **Write values into a staged file before the application opens it.** Nothing should depend on a
  later save happening.
- **A template's extension is not the document's.** A staged `.dotx`/`.xltx`/`.potx` must become
  `.docx`/`.xlsx`/`.pptx` for the staged copy *and* the vault dataset. `StagedFileName` in the
  service does this for every host; a template staged under its own extension made LibreOffice open
  "Untitled 1".
- **Datasets are named by their part number** — the service already does this for every host.
- **One typing rule, shared.** Do not add a second place that coerces a value for a field.
- **Never claim a `why` you have not measured.** Every "because ONLYOFFICE does X" in this repo
  must be something that was actually observed, with the version it was observed on.

## Decisions on the record

- **Editor plugin first** (Marc, Sep 24 2026), over a document-host/WOPI-style integration where
  Nexus serves files to Document Server and takes the save callback. That remains a reasonable
  second deliverable; it is server work, not an add-in, and would belong in its own repo or in the
  Engine.
- **Named `.Addins`** to match `Nexus.PLM.LibreOffice.Addins` and `Nexus.PLM.Office.Addins`, even
  though ONLYOFFICE's own word is "plugin".

## Tests

```bash
npm test          # node --test, no install, no framework, no editor
```

Everything that can be decided without an editor belongs in `plugin/lib/` where the tests reach
it — which commands exist and when (`commands.js`), the tab payload (`toolbar.js`), what to do
with an answer (`host.js`), comments to markups and back (`markup.js`), where a file is
(`paths.js`), the panel (`panel.js`). `editor.js` (the editor and the desktop shell), `runner.js`
(the one path a press takes, shared by the tab and the panel), `background.js`, `ui.js` and
`plugin.js` are the glue that cannot be tested that way and are deliberately dull.

## Branches

`main` and `next` are protected; work goes on a `Marc/` branch and PRs target **`next`**. Never
commit to `main` or `next` directly.
