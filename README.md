# Nexus PLM for ONLYOFFICE

[![Build](https://github.com/Nexus-PLM/Nexus.PLM.OnlyOffice.Addins/actions/workflows/build.yml/badge.svg)](https://github.com/Nexus-PLM/Nexus.PLM.OnlyOffice.Addins/actions/workflows/build.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

Product lifecycle management from inside the ONLYOFFICE editors. Check a document out, see what
PLM knows about it, edit its attributes and check it back in — in **Document**, **Spreadsheet**
and **Presentation**, without leaving the editor.

> **Status: the Nexus PLM tab is in the ribbon and runs.** Word's tab — its seven groups, all 26
> commands, split-button menus and icons — sits in the ribbon of all three editors in ONLYOFFICE
> Desktop Editors 9.4.0, and the buttons do what they do in Word through the same Addin Service:
> Sign In, Open (into a new tab), Check Out, Check In (the document is saved first), Edit Values,
> Refresh Values and Connection Status have been driven end to end; the rest post the same bodies
> Word does and are covered by tests rather than by a hand on the mouse yet. A docked **Navigator**
> browses the vault beside the document — the same folder tree, rows and four actions as Word's
> pane — and it has been driven end to end too: Open, Check Out, Check In, Properties, search and
> Refresh.

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
  config.json        three variations, every one for all three editors:
                     the resident background half (the tab), the Navigator panel, About
  background.html/js the tab: drawn once, updated as the state changes, clicks handed to the runner
  index.html, ui.js  the Navigator pane's frame and DOM, and nothing it decides for itself
  runner.js          the one path a press takes, whether from the tab or the panel
  editor.js          the editor and the desktop shell, as a plugin frame can reach them
  plugin.js          the Asc.plugin callbacks for the panel, and nothing else
  lib/
    commands.js      the ONE table: every command, its group, endpoint, body and enable rule
    toolbar.js       the table as the AddToolbarMenuItem payload
    host.js          the host's half of each command, as data (open this, write these values)
    markup.js        comments to review markups and back
    paths.js         where a staged file is, remembered by name
    navigator.js     the Navigator's model: the tree, and what may be done to the row you pick
    panel.js         what the pane says when it cannot show the vault - no ONLYOFFICE, no DOM
    editors.js       the three editors and the only places they differ
    client.js        talking to the Nexus PLM Addin Service
  icons/             light and dark, five scales each, from the add-in family's own art
tools/
  install.ps1        copy into ONLYOFFICE's user plugin folder and write the secret
  make-icons.py      regenerate the icon set from the 80px originals
tests/               node --test, no editor and no browser needed
```

The measured facts behind the tab, the file and the shell — what works from a plugin frame and
what only looked like it should — are in `CLAUDE.md`, so that they are read before anything here
is changed.

## The Navigator

Docked beside the document, the same pane Word carries: a folder tree on top, the chosen folder's
documents below, and four buttons that act on the row you pick — Open, Check Out, Check In,
Properties. It is the vault, not the open file: everything it shows and does is about the selected
row, so Properties on a search hit opens that hit, not the document you happen to be editing.

The tree, the rows, the band heading and which buttons are live are all decided in
`plugin/lib/navigator.js`, where the tests reach them without an editor and without a browser.
`ui.js` only draws what it is told. The search box waits for two characters and 300 ms before it
asks the service, which are Word's numbers.

**One rule differs from Word's pane, and deliberately.** Check In needs your own lock *and* that
document open in this tab, because the service uploads a working copy and this host only has a path
for the document its frame is docked to. Word's pane enables the button on the lock alone and then
refuses the press with a toast — "is checked out to you but is not open here. Open it first, then
check it in." Here that same sentence is the disabled button's tooltip, so it arrives before the
click is spent rather than after.

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
