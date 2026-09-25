# Nexus PLM for ONLYOFFICE

[![Build](https://github.com/Nexus-PLM/Nexus.PLM.OnlyOffice.Addins/actions/workflows/build.yml/badge.svg)](https://github.com/Nexus-PLM/Nexus.PLM.OnlyOffice.Addins/actions/workflows/build.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

Product lifecycle management from inside the ONLYOFFICE editors. Check a document out, see what
PLM knows about it, edit its attributes and check it back in — in **Document**, **Spreadsheet**
and **Presentation**, without leaving the editor.

> **Status: early, but no longer unproven.** The plugin installs, registers, and appears in the
> Plugins ribbon of all three editors in ONLYOFFICE Desktop Editors 9.4.0. The commands are not
> wired up, because a measurement says they cannot work yet — see
> [What a plugin frame can and cannot do](#what-a-plugin-frame-can-and-cannot-do--measured-desktop-editors-940-sep-24-2026).

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

## What a plugin frame can and cannot do — measured, Desktop Editors 9.4.0, Sep 24 2026

The plugin is installed, it registers, it appears in the **Plugins** ribbon in Document,
Spreadsheet and Presentation, and `Asc.plugin` is present. A probe build reported, from inside a
real editor with a real document open:

```
editorType    = "word"
page origin   = file://
page protocol = file:
Asc.plugin    = present
fetch         = function
FETCH FAILED: Failed to fetch
XHR FAILED   (blocked or unreachable)
```

**It is CORS, and only CORS.** The two candidates were separated rather than guessed between:

- **Not unreachable.** `GET http://localhost:5100/api/health` with no `Origin` header answers
  `200 {"success":true,...}` from the same machine at the same moment.
- **Not mixed content.** The plugin page is served as `file:`, not `https:`, so a browser permits
  `http://` from it. Mixed content is a problem for Docs in a browser, not for Desktop Editors.
- **CORS.** A `file://` page has an opaque origin, so Chromium sends `Origin: null` and applies
  CORS. The service answers `200` — and sends **no `Access-Control-Allow-Origin` at all**, on any
  origin (`null`, `file://`, `https://...`). The `OPTIONS` preflight answers `204` with no
  `Access-Control-Allow-Origin`, `-Allow-Methods` or `-Allow-Headers`. So the browser makes the
  call, gets the answer, and throws it away.

### What that means

**The AddinService needs CORS headers before any command in this plugin can work.** That is a real
service change, and it does not contradict "the service needs no changes for a new host": that rule
is about commands and dialogs — a host declares what it can do and the behaviour follows. CORS is
not a host capability, it is a property of *every* browser-hosted client, and the service has never
had one before. This one is owed.

What is needed: `Access-Control-Allow-Origin` covering the opaque origin a `file://` plugin sends
(`null`, or `*`), and a preflight answer carrying the methods and headers the plugin uses. The
plugin sends no cookies, so the credentials restriction on `*` does not bite.

### Still unmeasured

**Docs in a browser, over https.** There the plugin page is `https:` and `http://localhost:5100`
is mixed content, which CORS headers alone will **not** fix — the browser refuses the request
before any header is read. That is a prediction from how the two mechanisms work, not a
measurement, and it has not been tested: nothing was listening on 80, 443, 8080 or 8443 on the
server when this was written. Settle it before assuming one fix serves both.

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
