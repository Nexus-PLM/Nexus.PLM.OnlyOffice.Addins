# Nexus PLM for ONLYOFFICE

[![Build](https://github.com/Nexus-PLM/Nexus.PLM.OnlyOffice.Addins/actions/workflows/build.yml/badge.svg)](https://github.com/Nexus-PLM/Nexus.PLM.OnlyOffice.Addins/actions/workflows/build.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

Product lifecycle management from inside the ONLYOFFICE editors. Check a document out, see what
PLM knows about it, edit its attributes and check it back in — in **Document**, **Spreadsheet**
and **Presentation**, without leaving the editor.

> **Status: early.** The plugin loads, reports which editor it is in, and draws the panel. The
> commands are not wired up yet, and one thing has to be measured before they can be — see
> [The open question](#the-open-question). Nothing here has been run against a real ONLYOFFICE.

---

## The shape of it

One plugin serves all three editors. `config.json` declares
`EditorsSupport: ["word", "cell", "slide"]` in a **single variation**, and nothing in the code is
editor-specific: the plugin asks ONLYOFFICE which editor it is in once, and the few places that
genuinely differ live in `plugin/lib/editors.js`.

That is deliberate, and it is the architecture this repo exists to hold. The Word, Excel and
PowerPoint add-ins are three codebases, and they drifted — one of them blanked typed fields that an
empty value should have left alone, because only the other had been fixed. Three variations here
would be three plugins in a trench coat and would drift the same way. A test asserts there is
exactly one.

```
plugin/
  config.json        one variation, three editors
  index.html         the panel's frame
  plugin.js          the Asc.plugin callbacks, and nothing else
  ui.js              puts decisions into the DOM
  lib/
    editors.js       the three editors and the only places they differ
    panel.js         what the panel shows - no ONLYOFFICE, no DOM
    client.js        talking to the Nexus PLM Addin Service
tests/               node --test, no editor and no browser needed
```

## It needs the Addin Service

The plugin talks to the **Nexus PLM Addin Service** on `localhost:5100`, which owns the dialogs and
does the talking to the PLM server — so the same windows, wording and behaviour appear in
ONLYOFFICE, LibreOffice, Word, Excel, PowerPoint and FreeCAD. A host declares what it can do and it
travels with the request; **the service needs no change for a new host.** If this plugin ever seems
to need one, the question to ask is what it should be declaring instead.

## The open question

Every other Nexus host is a desktop process. This one is JavaScript in a browser frame, and two
things stand between it and the service. Both must be measured before the commands are built:

1. **CORS.** The service answers desktop clients, which send no `Origin`. A browser sends one and
   will discard the response without `Access-Control-Allow-Origin`.
2. **Mixed content.** A browser on an `https://` page will not call `http://localhost:5100` at all,
   whatever the service allows.

Neither changes where the dialogs belong. But until both are measured against a real ONLYOFFICE —
Desktop Editors and Docs in a browser are likely to differ — nothing that assumes the call
succeeds should be written.

## No connector in this repo, on purpose

ONLYOFFICE reads and writes **OOXML** — the same `.docx`, `.xlsx` and `.pptx` the Office add-ins
use. `Nexus.PLM.Office.*.Templates` already discovers, reads and writes the drivable fields of
those files through OpenXml only, with no Office installed and no COM, and the Nexus Vault already
runs them server-side. So there is nothing for this repo to add there, and adding it would mean two
implementations of one format.

## Tests

```bash
npm test
```

No install, no framework, no editor: Node's own test runner over the plain modules under
`plugin/lib/`. Everything that can be decided without ONLYOFFICE is decided there so that it can be
tested; what is left in `plugin.js` and `ui.js` is glue that cannot.

## Licence

MIT. See [LICENSE](LICENSE).
