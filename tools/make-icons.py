"""
The Nexus PLM icon set for the ONLYOFFICE tab.

The art is the add-in family's own: the 80px originals the Word ribbon and the LibreOffice toolbar
are drawn from, so a command looks the same in every host. ONLYOFFICE asks for a big button's
icon at five display scales and names the files for them itself - `x.png`, `x@1.25x.png`,
`x@1.5x.png`, `x@1.75x.png`, `x@2x.png` - so that is what is written, at the 28px base the
editor's own big buttons use (measured against the AI plugin shipped with Desktop Editors 9.4.0).

The dark set is the same art: grey strokes with one teal accent read on either ground, and the
LibreOffice add-in ships a single set for the same reason.

    python tools/make-icons.py [--source <dir of *-80x80.png>]

Needs Pillow. The default source is the LibreOffice add-in's `extension/icons/src` in a sibling
checkout; pass --source to point anywhere else.
"""

import argparse
import os
import sys

from PIL import Image

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "..", "plugin", "icons")
DEFAULT_SOURCE = os.path.join(HERE, "..", "..", "Nexus.PLM.LibreOffice.Addins", "extension", "icons", "src")

# Command id -> the original's name. Every command on the tab, plus the plugin's own icon.
ART = {
    "sign_in": "SignIn", "sign_out": "SignOut",
    "new": "New", "open": "Open", "save": "Save", "save_as_new": "SaveAs",
    "save_as_existing": "SaveAsExisting",
    "navigator": "Navigate", "search": "Search", "properties": "Properties",
    "check_out": "CheckOut", "check_in": "CheckIn", "release": "Release", "revise": "Revise",
    "markup": "MarkUp", "apply_markups": "MarkUp", "change_owner": "ChangeOwner",
    "worklist": "MyWorkList", "new_workflow": "NewWorkflow",
    "edit_values": "EditValues", "refresh_values": "RefreshValues", "reload_document": "Reload",
    "settings": "Settings", "help": "Help", "about": "Nexus", "connection": "ConnectionStatus",
    "icon": "Nexus",
}

# Display scale -> (pixel size, file suffix). 28px is the editor's own big-button size at 100%.
SCALES = [(100, 28, ""), (125, 35, "@1.25x"), (150, 42, "@1.5x"), (175, 49, "@1.75x"), (200, 56, "@2x")]


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--source", default=DEFAULT_SOURCE)
    args = parser.parse_args()

    missing = [n for n in set(ART.values()) if not os.path.exists(os.path.join(args.source, n + "-80x80.png"))]
    if missing:
        sys.exit("missing originals in %s: %s" % (args.source, ", ".join(sorted(missing))))

    for theme in ("light", "dark"):
        folder = os.path.join(OUT, theme)
        os.makedirs(folder, exist_ok=True)
        # Start clean: a stale file at a size nothing asks for is a file nobody notices.
        for old in os.listdir(folder):
            if old.endswith(".png"):
                os.remove(os.path.join(folder, old))
        for command, name in ART.items():
            original = Image.open(os.path.join(args.source, name + "-80x80.png")).convert("RGBA")
            for _, size, suffix in SCALES:
                original.resize((size, size), Image.LANCZOS).save(os.path.join(folder, command + suffix + ".png"))
    print("wrote %d icons per theme" % (len(ART) * len(SCALES)))


if __name__ == "__main__":
    main()
