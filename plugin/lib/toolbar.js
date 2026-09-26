/*
 * The Nexus PLM tab, as ONLYOFFICE wants it described.
 *
 * `Asc.plugin.executeMethod("AddToolbarMenuItem", [payload])` adds a named tab to the ribbon with
 * buttons in it, and `UpdateToolbarMenuItem` with the same payload brings the buttons it already
 * drew up to date. This file turns the one command table into that payload and nothing else — no
 * decision about which commands exist or when they may be pressed is taken here, so the tab and
 * the panel cannot drift apart.
 *
 * ── What the editor reads from an item (measured, Desktop Editors 9.4.0) ─────────────────────
 * Read from web-apps' `Common.UI.LayoutManager.addCustomControls`, the function that draws it:
 *   id, text, hint, icons        the button
 *   type: "big-button"           the large icon-over-caption button Word's ribbon uses throughout
 *   separator: true              a separator BEFORE this button, and a new group — Word's groups
 *   split: true + items: [...]   a split button whose arrow opens a menu of {id, text} entries;
 *                                a menu entry's click arrives under its own id, like a button's
 *   enableToggle: true           a button that stays pressed
 *   disabled: true               greyed, menu and all — which is why nothing in a menu may be
 *                                enabled while its button is not (commands.js keeps it so)
 *   lockInViewMode: true         greyed while the document is read-only
 * On an UPDATE the editor changes an existing button's text, hint, disabled state and menu, and
 * nothing else: an icon cannot change after the tab is drawn.
 *
 * Built without an editor so the tests can read it. The only thing that needs a real ONLYOFFICE is
 * handing the result to executeMethod, which `background.js` does.
 */

(function (root, factory) {
    if (typeof module === "object" && module.exports) {
        module.exports = factory(require("./commands.js"));
    } else {
        root.NexusPlmToolbar = factory(root.NexusPlmCommands);
    }
}(typeof self !== "undefined" ? self : this, function (commands) {
    "use strict";

    /** Prefix on every button id, so nothing of ours can collide with the editor's own. */
    var PREFIX = "nexusplm_";

    /** The tab's own id. */
    var TAB_ID = PREFIX + "tab";

    /** Every button on the tab is the large kind, as every button on Word's ribbon is. */
    var BIG_BUTTON = "big-button";

    function buttonId(commandId) { return PREFIX + commandId; }

    /**
     * Where a command's icon lives, as ONLYOFFICE wants it: a PATTERN, not a path.
     *
     * The editor fills in the placeholders itself — the theme it is drawing in and the display
     * scale — and then looks for that file. `%scale%(default)` stands for the five scales it
     * knows, 100% to 200%, and it names the files for them itself:
     *
     *     icons/light/check_out.png   check_out@1.25x.png   @1.5x   @1.75x   check_out@2x.png
     *
     * (Read from web-apps' `Common.UI.iconsStr2IconsObj`.) One string per command therefore
     * covers light and dark and every scale, and `tools/make-icons.py` writes exactly those
     * files from the add-in family's own art, so a command looks the same in every host.
     */
    function iconsFor(command) {
        return "icons/%theme-type%(light|dark)/" + command.id + "%scale%(default).png";
    }

    /** The command a toolbar button or menu entry id refers to, or null if the id is not ours. */
    function commandFor(buttonId) {
        if (typeof buttonId !== "string" || buttonId.indexOf(PREFIX) !== 0) { return null; }
        return commands.byId(buttonId.slice(PREFIX.length));
    }

    /** The tooltip: the reason when a command is greyed, so a disabled button explains itself. */
    function hintFor(command, context) {
        return commands.disabledBecause(command, context) || command.label;
    }

    /** One entry in a split button's menu. */
    function menuEntry(command, context) {
        return {
            id: buttonId(command.id),
            text: command.label,
            hint: hintFor(command, context),
            disabled: !commands.isEnabled(command, context)
        };
    }

    /**
     * One button, with its menu when the table gives it one.
     *
     * `first` is whether it opens a new group: the editor draws the separator before the button
     * that carries the flag, so the first button of every group but the first carries it.
     */
    function button(command, context, first) {
        var item = {
            id: buttonId(command.id),
            type: BIG_BUTTON,
            text: command.label,
            hint: hintFor(command, context),
            icons: iconsFor(command),
            disabled: !commands.isEnabled(command, context),
            lockInViewMode: true
        };
        if (first) { item.separator = true; }
        if (command.toggle) { item.enableToggle = true; }

        var children = commands.childrenOf(command.id);
        if (children.length) {
            item.split = true;
            item.items = children.map(function (child) { return menuEntry(child, context); });
        }
        return item;
    }

    /**
     * The whole payload for AddToolbarMenuItem and UpdateToolbarMenuItem.
     *
     * `guid` must be the plugin's own — the editor uses it to know whose tab this is. The groups
     * are Word's, in Word's order, each opened by a separator.
     */
    function tab(guid, context) {
        var items = [];
        commands.GROUPS.forEach(function (group, g) {
            commands.inGroup(group).filter(function (c) { return !c.parent; })
                .forEach(function (command, i) {
                    items.push(button(command, context, g > 0 && i === 0));
                });
        });

        return {
            guid: guid,
            tabs: [{
                id: TAB_ID,
                text: commands.TAB,
                items: items
            }]
        };
    }

    /** Every id a click can arrive under: the buttons and the entries in their menus. */
    function clickIds() {
        return commands.COMMANDS.map(function (command) { return buttonId(command.id); });
    }

    /** Every button's id and disabled state, the summary the tests compare states with. */
    function buttonStates(context) {
        return commands.COMMANDS.map(function (command) {
            return {
                id: buttonId(command.id),
                disabled: !commands.isEnabled(command, context)
            };
        });
    }

    return {
        PREFIX: PREFIX,
        TAB_ID: TAB_ID,
        BIG_BUTTON: BIG_BUTTON,
        buttonId: buttonId,
        iconsFor: iconsFor,
        commandFor: commandFor,
        button: button,
        tab: tab,
        clickIds: clickIds,
        buttonStates: buttonStates
    };
}));
