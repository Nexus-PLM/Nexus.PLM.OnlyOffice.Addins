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

**One plugin, three editors.** `config.json` has a **single variation** declaring
`EditorsSupport: ["word", "cell", "slide"]`. Nothing outside `lib/editors.js` may branch on which
editor it is in. A test asserts there is exactly one variation and that `config.json` and the code
agree about which editors exist — because neither disagreement fails loudly inside ONLYOFFICE.

## The open question — measure this before building commands

This is the first Nexus host that is **not a desktop process**. The plugin is JavaScript in a
browser frame, and two things stand between it and the Addin Service on `localhost:5100`:

1. **CORS.** The service answers desktop clients, which send no `Origin`. A browser sends one and
   discards the response without `Access-Control-Allow-Origin`.
2. **Mixed content.** A browser on an `https://` page will not call `http://localhost:5100` at all.

Measure both, on **Desktop Editors and on Docs in a browser** — they may well differ. Write down
what was measured. Do not design around either until then, and do not assume a service change is
the answer: the rule below still applies.

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
