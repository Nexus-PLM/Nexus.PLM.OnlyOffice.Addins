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

## The next slice: which file the editor has open

The panel currently reports every document as not in PLM because it does not know the file. What
is known so far, from ONLYOFFICE's own docs and source rather than from guessing:

- **`Asc.plugin.info` does not carry it.** Its documented fields are `data`, `editorType`, `guid`,
  `height`, `imgSrc`, `mmToPx`, `objectId`, `recalculate`, `resize`, `width`. No name, path or URL.
- **The lead worth following is `initDataType: "desktop-external"`** — "the main page data of the
  desktop app (system messages)". It is what ONLYOFFICE's own **encryption** plugins declare, and
  those work against local documents; there is a doc page, `desktop-editors/get-started/
  how-it-works/encrypting-local-documents`. A second variation can declare it, since variations do
  not all have to be alike — only all three editors have to be served.

Do not invent a mechanism here. Read the encryption plugins first: they are the ones already
solving "a plugin that needs to know about the local file".

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
it — which rows, which buttons are live, what to say when PLM has never seen the document.
`plugin.js` and `ui.js` are the glue that cannot be tested that way and are deliberately dull.

## Branches

`main` and `next` are protected; work goes on a `Marc/` branch and PRs target **`next`**. Never
commit to `main` or `next` directly.
