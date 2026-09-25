/*
 * The three editors, and the only places they differ.
 *
 * ONLYOFFICE names them "word", "cell" and "slide" — Document, Spreadsheet and Presentation as a
 * user sees them. The plugin is one codebase for all three, so every difference belongs here
 * where it can be read at a glance and tested without an editor.
 *
 * Plain UMD so the same file loads in the editor's sandbox (a <script> tag, no bundler) and in
 * the tests (a require). The editor's plugin frame has no module loader.
 */

(function (root, factory) {
    if (typeof module === "object" && module.exports) { module.exports = factory(); }
    else { root.NexusPlmEditors = factory(); }
}(typeof self !== "undefined" ? self : this, function () {
    "use strict";

    var WORD = "word";
    var CELL = "cell";
    var SLIDE = "slide";

    var EDITORS = {};

    EDITORS[WORD] = {
        key: WORD,
        //: What a user calls it. Used in messages, so it must match ONLYOFFICE's own wording.
        name: "Document",
        //: The extension a new item of this kind gets. OOXML, the same formats the Office
        //: add-ins use — which is why the server-side field connectors already understand them
        //: and this repo needs no connector of its own.
        extension: ".docx",
        templateExtension: ".dotx",
        //: Where a template author can put a drivable field, and what the connector calls it.
        //: A spreadsheet's is a defined name over exactly one cell; a document's is a content
        //: control or a custom property; a presentation's is a named shape.
        fieldKinds: ["custom property", "content control"]
    };

    EDITORS[CELL] = {
        key: CELL,
        name: "Spreadsheet",
        extension: ".xlsx",
        templateExtension: ".xltx",
        fieldKinds: ["custom property", "defined name"]
    };

    EDITORS[SLIDE] = {
        key: SLIDE,
        name: "Presentation",
        extension: ".pptx",
        templateExtension: ".potx",
        fieldKinds: ["custom property", "named shape"]
    };

    /**
     * What we know about the editor the plugin is running in.
     *
     * An unknown editor is not an error: ONLYOFFICE may add one, and a plugin that refuses to
     * draw is worse than one that shows the item and hides the few commands it cannot place.
     * The caller gets null and shows the document-less panel.
     */
    function describe(editorType) {
        if (!editorType) { return null; }
        return EDITORS[String(editorType).toLowerCase()] || null;
    }

    /** Every editor this plugin declares support for, in config.json order. */
    function supported() {
        return [EDITORS[WORD], EDITORS[CELL], EDITORS[SLIDE]];
    }

    return {
        WORD: WORD, CELL: CELL, SLIDE: SLIDE,
        describe: describe,
        supported: supported
    };
}));
