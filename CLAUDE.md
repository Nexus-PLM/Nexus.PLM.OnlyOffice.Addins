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

## Measured: a plugin frame cannot reach the service yet (Desktop Editors 9.4.0, Sep 24 2026)

The plugin registers and appears in the Plugins ribbon in all three editors. From inside a real
editor, `fetch` and `XMLHttpRequest` to `http://localhost:5100` **both fail**, while the identical
request from outside the browser succeeds — so the service is up and the call is being blocked.

It is **CORS**, established by separating the candidates, not by guessing:

- The plugin page's origin is `file://` and its protocol `file:` — so this is **not** mixed
  content, and a browser permits `http://` from it.
- The service answers `200`/`204` but sends **no** `Access-Control-Allow-Origin` on any origin
  (`null`, `file://`, `https://...`), and no preflight headers either.

**The AddinService owes CORS headers.** This does not contradict "the service needs no changes for
a new host": that rule is about commands and dialogs. CORS is a property of every browser-hosted
client, and no host before this one was a browser.

**Still unmeasured:** Docs in a browser over https, where the page is `https:` and
`http://localhost:5100` is mixed content — which CORS headers alone will **not** fix. That is a
prediction, not a measurement. Nothing was listening on 80/443/8080/8443 on the server when this
was written.

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
