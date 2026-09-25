# Nexus PLM for ONLYOFFICE

[![Build](https://github.com/Nexus-PLM/Nexus.PLM.OnlyOffice.Addins/actions/workflows/build.yml/badge.svg)](https://github.com/Nexus-PLM/Nexus.PLM.OnlyOffice.Addins/actions/workflows/build.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

Product lifecycle management from inside the ONLYOFFICE editors. Check a document out, see what
PLM knows about it, edit its attributes and check it back in — in **Document**, **Spreadsheet**
and **Presentation**, without leaving the editor.

> **Status: early, and talking to PLM.** The plugin installs, registers, appears in the Plugins
> ribbon of all three editors in ONLYOFFICE Desktop Editors 9.4.0, reaches the Addin Service, and
> reports the signed-in user. The commands themselves are the next slice: the panel does not yet
> know which file the editor has open, so it honestly reports the document as not in PLM.

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

## How it reaches the service — measured, Desktop Editors 9.4.0, Sep 24 2026

The plugin declares **`"onlyofficeScheme": true`**, and that one line is what makes the whole thing
possible. With it, the editor serves the plugin under its own registered scheme and
`window.location.origin` is `onlyoffice://plugin`. Without it the page is `file://`, its origin is
*opaque*, and **every** `fetch` and `XMLHttpRequest` to the Addin Service fails — measured, before
and after.

That was not guessed at. Of the plugins shipped with Desktop Editors, exactly three make network
calls — AI, AI agent and DeepL — and exactly those three declare the flag. `sdk-all.js` rewrites
their base URL to `onlyoffice://plugin/<path>`.

It matters beyond getting a call through: a named scheme registered by the editor **cannot be
claimed by a web page**, whereas an opaque origin can be obtained by any website through a
sandboxed iframe. So the AddinService allows `onlyoffice://plugin` specifically, and does not have
to allow `null` at all.

### Still unmeasured

**Docs in a browser, over https.** There the page would be `https:` and `http://localhost:5100`
is mixed content, which no CORS or scheme change fixes — the browser refuses before any header is
read. That is a prediction from how the mechanisms work, not a measurement; nothing was listening
on 80/443/8080/8443 on the server when this was written.

## No connector in this repo, on purpose

ONLYOFFICE reads and writes **OOXML** — the same `.docx`, `.xlsx` and `.pptx` the Office add-ins
use. `Nexus.PLM.Office.*.Templates` already discovers, reads and writes the drivable fields of
those files through OpenXml only, with no Office installed and no COM, and the Nexus Vault already
runs them server-side. So there is nothing for this repo to add there, and adding it would mean two
implementations of one format.

## Installing it

```powershell
pwsh tools/install.ps1
```

Copies the plugin into ONLYOFFICE's **user** plugin folder — the one under Program Files needs
elevation, and the user folder has its own `v1`, so the `../v1/plugins.js` the page loads still
resolves.

## Tests

```bash
npm test
```

No install, no framework, no editor: Node's own test runner over the plain modules under
`plugin/lib/`. Everything that can be decided without ONLYOFFICE is decided there so that it can be
tested; what is left in `plugin.js` and `ui.js` is glue that cannot.

## Licence

MIT. See [LICENSE](LICENSE).
