/*
 * The Nexus PLM tab, as ONLYOFFICE wants it described.
 *
 * `Asc.plugin.executeMethod("AddToolbarMenuItem", [payload])` adds a named tab to the ribbon with
 * buttons in it. This file turns the one command table into that payload and nothing else — no
 * decision about which commands exist or when they may be pressed is taken here, so the tab and
 * the panel cannot drift apart.
 *
 * Built without an editor so the tests can read it. The only thing that needs a real ONLYOFFICE is
 * handing the result to executeMethod, which `plugin.js` does.
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

    function buttonId(commandId) { return PREFIX + commandId; }

    /**
     * Where a command's icon lives, as ONLYOFFICE wants it: a PATTERN, not a path.
     *
     * The editor fills in the placeholders itself — the theme it is drawing in, the button state,
     * and the display scale — and then looks for that file. So one string per command covers
     * light and dark and every scale, and the files simply have to be named to match:
     *
     *     icons/light/check_out.png        icons/light/check_out@2x.png
     *     icons/dark/check_out.png         icons/dark/check_out@2x.png
     *
     * The art is the add-in's own, the same icons the LibreOffice toolbar and the Word ribbon
     * use, so a command looks the same in every host.
     */
    function iconsFor(command) {
        return "icons/%theme-type%(light|dark)/" + command.id +
               "%state%(normal)%scale%(default|*).png";
    }

    /** The command a toolbar button id refers to, or null if the id is not one of ours. */
    function commandFor(buttonId) {
        if (typeof buttonId !== "string" || buttonId.indexOf(PREFIX) !== 0) { return null; }
        return commands.byId(buttonId.slice(PREFIX.length));
    }

    /**
     * One button.
     *
     * The tooltip carries the reason when a command is greyed, so a disabled button explains
     * itself instead of being a dead end the user pokes at.
     */
    function button(command, context) {
        var why = commands.disabledBecause(command, context);
        var item = {
            id: buttonId(command.id),
            type: "button",
            text: command.label,
            hint: why || command.label,
            icons: iconsFor(command),
            disabled: !commands.isEnabled(command, context),
            lockInViewMode: true
        };
        if (command.toggle) { item.enableToggle = true; }
        return item;
    }

    /**
     * The whole payload for AddToolbarMenuItem.
     *
     * `guid` must be the plugin's own — the editor uses it to know whose tab this is, and to
     * replace the tab rather than add a second one when it is called again.
     */
    function tab(guid, context) {
        var items = [];
        commands.GROUPS.forEach(function (group) {
            commands.inGroup(group).forEach(function (command) {
                items.push(button(command, context));
            });
        });

        return {
            guid: guid,
            tabs: [{
                id: PREFIX + "tab",
                text: commands.TAB,
                items: items
            }]
        };
    }

    /**
     * Every button's id and disabled state, for refreshing the tab after the state changes.
     *
     * The whole tab is re-sent rather than patched: AddToolbarMenuItem replaces the tab for this
     * guid, and one call that is always right beats a diff that is usually right.
     */
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
        buttonId: buttonId,
        iconsFor: iconsFor,
        commandFor: commandFor,
        button: button,
        tab: tab,
        buttonStates: buttonStates
    };
}));
