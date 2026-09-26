/*
 * About: the service's own window, not one of ours.
 *
 * Every Nexus host shows the same About - version, the server it is talking to, who is signed in -
 * because the service owns it. A second one written here would be a second thing to keep true.
 */

(function (window) {
    "use strict";

    window.Asc = window.Asc || {};
    window.Asc.plugin = window.Asc.plugin || {};

    window.Asc.plugin.init = function () {
        var client = new window.NexusPlmClient.Client();
        client.command("/plm/about", {
            hwnd: 0,
            host_name: "ONLYOFFICE",
            addin_version: "0.2.0"
        }).then(function () {
            try { window.Asc.plugin.executeCommand("close", ""); } catch (e) { /* closing */ }
        });
    };

    window.Asc.plugin.button = function () {};
})(window);
