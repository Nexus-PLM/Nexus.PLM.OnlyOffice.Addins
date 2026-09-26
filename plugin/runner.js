/*
 * Running a command: the one path a press takes, whether it came from the tab or the panel.
 *
 * Both frames — the resident background half that owns the tab, and the Navigator panel — use
 * this same runner, so a command cannot behave differently depending on where it was pressed.
 * Each frame makes one, gives it the editor it is in, and is told whenever the state changes so
 * it can redraw whatever it draws.
 *
 * Every decision is taken in lib/ and tested there: which commands exist and when each may be
 * pressed (commands.js), what to do with an answer (host.js), how comments become markups
 * (markup.js). What is left here is the order things happen in, and the calls that need a real
 * editor (editor.js).
 */

(function (window) {
    "use strict";

    var commands = window.NexusPlmCommands;
    var host = window.NexusPlmHost;
    var markup = window.NexusPlmMarkup;
    var clientModule = window.NexusPlmClient;
    var editor = window.NexusPlmEditor;

    /** Where Help goes: the product site. */
    var HELP_URL = "https://nexus-plm.com/help";

    /**
     * @param options.editorType  "word" | "cell" | "slide"
     * @param options.onState     called with the context after every read of the state
     * @param options.client      the service client (a test can hand in its own)
     */
    function create(options) {
        var o = options || {};
        var client = o.client || new clientModule.Client();
        var editorType = o.editorType || null;
        var onState = o.onState || function () {};

        /** Who is signed in, and what PLM says the open file is. The tab and panel draw from it. */
        var context = { signedIn: false, user: null, state: null, trouble: {}, doc: null };

        /**
         * The item this document is known to be, once anything has said so. Read by item from
         * then on, as Word does after Open: a staged path never matches what the Engine
         * recorded, and a Save As registers the file before the path could resolve it.
         */
        var bound = null;

        /** Template fields seen for this document, so Edit Values can read them back. */
        var fieldNames = [];

        /** Set while a command runs, so a second press does not start it twice. */
        var busy = false;

        // ── state ────────────────────────────────────────────────────────────

        /** Ask the service who is signed in and what the open document is. */
        function refresh(done) {
            var finish = function () { onState(context); if (done) { done(context); } };

            if (!clientModule.hasSecret()) {
                context.signedIn = false; context.user = null; context.state = null;
                context.trouble = { hasSecret: false };
                finish();
                return;
            }

            client.me().then(function (me) {
                // Three different failures, said three different ways. Reporting all of them as
                // "sign in" was wrong: only one of them is fixed by signing in.
                context.trouble = { unreachable: !!(me && me.unreachable), refused: !!(me && me.refused) };
                context.signedIn = !!(me && me.success && me.username);
                context.user = context.signedIn ? me.username : null;
                if (!context.signedIn) { context.state = null; finish(); return; }

                // By item once anything has said which; by path otherwise - the service resolves
                // a staged file from its name, which is the part number.
                editor.document(function (doc) {
                    context.doc = doc;
                    // The service resolves a file by its NAME (the part number), so the title
                    // is enough to ask with when the path is not known.
                    var query = bound && bound.itemId ? { itemId: bound.itemId } : { filePath: doc.path || doc.title };
                    if (!query.itemId && !query.filePath) { context.state = null; finish(); return; }

                    client.state(query).then(function (state) {
                        context.state = (state && state.success) ? state : null;
                        if (context.state && context.state.item_id && !bound) { bound = { itemId: context.state.item_id }; }
                        finish();
                    });
                });
            });
        }

        // ── running a command ────────────────────────────────────────────────

        function run(command) {
            if (!command || busy) { return; }

            // Signed out, a command that needs a session climbs the ladder instead of refusing -
            // what Word does, so one press signs in and then runs. `sign_in` itself is excluded:
            // it IS the ladder.
            if (!context.signedIn && commands.needsSession(command)) {
                ensureSignedIn(function (signedIn) {
                    if (signedIn) { run(command); }
                    else { say("Sign in to Nexus PLM to use " + command.label + ".", "info"); }
                });
                return;
            }

            // Never act on a command the rules say is unavailable. The tab greys it, but a stale
            // tab is one repaint away, and the service's refusal is indistinguishable from a bug.
            if (!commands.isEnabled(command, context)) {
                say(commands.disabledBecause(command, context));
                return;
            }

            if (command.id === "navigator") { showNavigator(); return; }
            if (command.id === "help") { openHelp(); return; }
            if (command.id === "connection") { reportConnection(); return; }

            busy = true;
            var finish = function () { busy = false; if (host.movesState(command)) { refresh(); } };

            withSavedFile(command, function (refusal) {
                if (refusal) { say(refusal, "warning"); busy = false; return; }

                if (command.id === "markup") { sendMarkups(command, finish); return; }
                if (command.id === "apply_markups") { applyMarkups(command, finish); return; }

                buildBody(command, function (body) {
                    client.command(command.endpoint, body).then(function (answer) {
                        carryOut(host.effectsFor(command, answer, context));
                        finish();
                    });
                });
            });
        }

        /**
         * Sign in, cheapest rung first, the way Word's SignInGate climbs: the session the service
         * already holds, then the one saved on this machine (silent), then the dialog. Answers
         * whether there is a session afterwards.
         */
        function ensureSignedIn(done) {
            busy = true;
            var finish = function (ok) { busy = false; refresh(function () { done(ok); }); };

            client.me().then(function (me) {
                if (me && me.success && me.username) { finish(true); return; }
                client.autoLogin().then(function (restored) {
                    if (restored && restored.success) { finish(true); return; }
                    client.command("/api/auth/login", { hwnd: 0 }).then(function (signedIn) {
                        finish(!!(signedIn && signedIn.success));
                    });
                });
            });
        }

        /**
         * A command that uploads the file needs the file current. Ask the editor to save first;
         * if it cannot, or the document still carries changes, refuse with the reason.
         */
        function withSavedFile(command, next) {
            if (!command.saves) { next(null); return; }
            editor.document(function (doc) {
                context.doc = doc;
                var refusal = host.refuse(command, doc);
                if (refusal) { next(refusal); return; }

                // Word saves the document before every upload and never asks. The editor does
                // not say whether there is anything to save, so it is asked to every time, and
                // given a moment to write before the file is read.
                editor.save(function () {
                    window.setTimeout(function () { next(null); }, 1500);
                });
            });
        }

        /** The body for this command, with the document's own values where the service wants them. */
        function buildBody(command, next) {
            var body = commands.bodyFor(command, context);
            if (command.id !== "edit_values" && command.id !== "save_as_new") { next(body); return; }

            // Edit Values shows what the document holds as the starting point; Save As pre-fills
            // the form from it. Word reads its custom properties and content controls for both.
            editor.readFields(editorType, fieldNames, function (values) {
                if (command.id === "edit_values") { body.document_values = values; }
                else { body.attributes = values; }
                next(body);
            });
        }

        /** Do what lib/host.js decided. */
        function carryOut(effects) {
            (effects || []).forEach(function (effect) {
                switch (effect.type) {
                    case host.OPEN:
                        // Remember where it is FIRST: the tab that opens it asks by name.
                        editor.rememberPath(effect.path);
                        editor.openFile(effect.path, function (asked) {
                            if (!asked) { say("Nexus PLM staged the file at " + effect.path + " but ONLYOFFICE could not open it from here.", "warning"); }
                        });
                        break;
                    case host.BIND:
                        bound = bound || {};
                        if (effect.itemId) { bound.itemId = effect.itemId; }
                        // The state is re-read from the service straight after; this only
                        // bridges the moment between the answer and that read.
                        if (context.state && effect.status) {
                            context.state.status = effect.status;
                            context.state.checked_out_by = effect.owner;
                            if (effect.revision) { context.state.revision = effect.revision; }
                            onState(context);
                        }
                        break;
                    case host.VALUES:
                        rememberFields(effect.values);
                        editor.writeFields(editorType, effect.values, function (written) {
                            var sent = Object.keys(effect.values || {}).length;
                            if (sent && !written) {
                                say("Nexus PLM sent " + sent + " value(s) but this document has no field that takes them.", "info");
                            }
                        });
                        break;
                    case host.SAY:
                        say(effect.message, effect.severity);
                        break;
                    default:
                        break;
                }
            });
        }

        function rememberFields(values) {
            Object.keys(values || {}).forEach(function (name) {
                if (fieldNames.indexOf(name) === -1) { fieldNames.push(name); }
            });
        }

        // ── markups ──────────────────────────────────────────────────────────

        function sendMarkups(command, finish) {
            editor.comments(function (list) {
                var entries = markup.toEntries(list);
                if (!entries.length) { say("No comments to send.", "info"); finish(); return; }
                var body = commands.bodyFor(command, context);
                body.entries = entries;
                client.command(command.endpoint, body).then(function (answer) {
                    carryOut(host.effectsFor(command, answer, context));
                    finish();
                });
            });
        }

        function applyMarkups(command, finish) {
            var itemId = context.state && context.state.item_id;
            client.markups(itemId).then(function (answer) {
                if (!answer || !answer.success) {
                    carryOut(host.effectsFor(command, answer, context));
                    finish();
                    return;
                }
                editor.comments(function (list) {
                    var additions = markup.planApply(answer.markups, list);
                    if (!additions.length) {
                        say("Every markup in Nexus PLM is already in this document.", "info");
                        finish();
                        return;
                    }
                    var left = additions.length;
                    additions.forEach(function (data) {
                        editor.addComment(data, function () {
                            if (--left === 0) {
                                say("Applied " + additions.length + " markup" + (additions.length === 1 ? "" : "s") + " from Nexus PLM.", "success");
                                finish();
                            }
                        });
                    });
                });
            });
        }

        // ── the local commands ───────────────────────────────────────────────

        function showNavigator() {
            var shown = editor.togglePanel();
            if (shown === null) {
                say("This build of ONLYOFFICE cannot open a panel from here. Use Plugins > Background plugins > Nexus PLM > Navigator.", "warning");
            }
        }

        function openHelp() {
            try { window.Asc.plugin.executeMethod("OpenLink", [HELP_URL]); }
            catch (e) { say("Could not open help.", "warning"); }
        }

        /**
         * Connection Status reports the service address and the session as a Nexus toast, which
         * doubles as the test: a toast the user can see is itself proof the service answered.
         * Word does the same.
         */
        function reportConnection() {
            client.health().then(function (answer) {
                if (answer && answer.unreachable) { say(answer.error, "error"); return; }
                editor.probe(function (seen) {
                    say("Connected to Nexus PLM on " + client.baseUrl + " - " +
                        (context.signedIn ? "signed in as " + context.user : "not signed in") + ". " + seen,
                        context.signedIn ? "success" : "warning");
                });
            });
        }

        /** Say something through the tray's own toast, so it looks like it does in every host. */
        function say(message, severity) {
            if (!message) { return; }
            client.notify(message, severity || "info");
        }

        return {
            context: context,
            client: client,
            refresh: refresh,
            run: run,
            runId: function (id) { run(commands.byId(id)); },
            say: say
        };
    }

    window.NexusPlmRunner = { create: create, HELP_URL: HELP_URL };

})(window);
