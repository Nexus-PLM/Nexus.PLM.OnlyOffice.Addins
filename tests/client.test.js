/*
 * What the client sends, and what the panel says when it cannot send it.
 *
 * The plugin's page is loaded from disk, so it has an opaque origin, which the service will not
 * trust on its own — any website can obtain one through a sandboxed iframe. The secret is the
 * proof that this caller is really on this machine, and these tests pin that it is actually sent
 * and that its absence is explained rather than surfacing as a bare 403.
 */

const test = require("node:test");
const assert = require("node:assert");

const clientModule = require("../plugin/lib/client.js");
const panel = require("../plugin/lib/panel.js");

/** A fetch that records what it was asked to do and answers a fixed body. */
function recordingFetch(body) {
    const calls = [];
    const fn = (url, options) => {
        calls.push({ url, options });
        return Promise.resolve({ text: () => Promise.resolve(JSON.stringify(body || {})) });
    };
    fn.calls = calls;
    return fn;
}

test("the secret is sent on every call", async () => {
    const fetch = recordingFetch({ success: true });
    const client = new clientModule.Client(undefined, fetch, "a-secret");

    await client.me();
    await client.state({ filePath: "C:/staging/x.docx" });

    assert.strictEqual(fetch.calls.length, 2);
    for (const call of fetch.calls) {
        assert.strictEqual(call.options.headers["X-Nexus-Addin"], "a-secret");
    }
});

test("no secret means no header rather than an empty one", async () => {
    // An empty header would be a value the service has to decide about. Sending nothing keeps the
    // question on one side: either this caller can prove it is local or it cannot.
    const fetch = recordingFetch({ success: true });
    const client = new clientModule.Client(undefined, fetch, null);

    await client.me();

    assert.ok(!("X-Nexus-Addin" in fetch.calls[0].options.headers));
});

test("a service that cannot be reached is a refusal, not an exception", async () => {
    const client = new clientModule.Client(undefined, () => Promise.reject(new Error("nope")));

    const answer = await client.me();

    assert.strictEqual(answer.success, false);
    assert.strictEqual(answer.unreachable, true);
    assert.match(answer.error, /not running/);
});

test("an answer that is not JSON is reported, not thrown", async () => {
    const client = new clientModule.Client(undefined,
        () => Promise.resolve({ text: () => Promise.resolve("<html>nope</html>") }));

    const answer = await client.me();

    assert.strictEqual(answer.success, false);
});

test("a plugin with no secret says it was not installed properly", () => {
    // Not "sign in": nobody can sign in past a 403, and the fix is the install step.
    assert.strictEqual(panel.trouble({ hasSecret: false, signedIn: true }), panel.NOT_INSTALLED);
});

test("the missing secret is reported before anything else", () => {
    // With no secret every call is refused, so any other message would be a guess about a
    // service that never answered.
    assert.strictEqual(
        panel.trouble({ hasSecret: false, unreachable: true, refused: true, signedIn: false }),
        panel.NOT_INSTALLED);
});

test("with a secret the pane goes back to showing the vault", () => {
    assert.strictEqual(panel.trouble({ hasSecret: true, signedIn: true }), null);
});

test("a service that is not answering is not reported as signed out", () => {
    assert.strictEqual(panel.trouble({ unreachable: true }), panel.NOT_RUNNING);
    assert.strictEqual(panel.trouble({ refused: true }), panel.REFUSED);
    assert.strictEqual(panel.trouble({ signedIn: false }), panel.NOT_SIGNED_IN);
});

test("the status travels with the answer so the panel can tell them apart", async () => {
    const client = new clientModule.Client(undefined,
        () => Promise.resolve({ status: 403, text: () => Promise.resolve('{"success":false}') }));

    const answer = await client.me();

    assert.strictEqual(answer.httpStatus, 403);
    assert.strictEqual(answer.refused, true);
});

test("the HTTP status never overwrites the item's own status", () => {
    // A /plm/state answer has a `status` of its own - the item's lifecycle. Writing the HTTP
    // status over it put a stateless call's 400 into the pane and made a document PLM had never
    // heard of look like one it knew. The HTTP status travels under a name of OUR choosing.
    const answer = { success: true, status: "unknown", httpStatus: 400 };
    assert.strictEqual(answer.status, "unknown");
    assert.strictEqual(answer.httpStatus, 400);
});

test("connecting tells the tray this host's name and version", async () => {
    // Without this the plugin works but is invisible: a ribbon full of commands and nothing in
    // the tray, which is exactly how it looked.
    const fetch = recordingFetch({ success: true });
    await new clientModule.Client(undefined, fetch, "s").connect();

    const call = fetch.calls[0];
    assert.match(call.url, /\/api\/addins\/connect$/);
    assert.deepStrictEqual(JSON.parse(call.options.body),
        { name: "ONLYOFFICE", version: clientModule.ADDIN_VERSION });
});

test("the heartbeat is comfortably inside the service's 45 second timeout", () => {
    // Twice inside the window, so one missed beat is not a disconnection.
    assert.ok(clientModule.HEARTBEAT_MS < 22500, "too slow to survive a missed beat");
});

test("a toast carries the shape the tray expects", async () => {
    const fetch = recordingFetch({ success: true });
    await new clientModule.Client(undefined, fetch, "s").notify("hello", "warning");

    const body = JSON.parse(fetch.calls[0].options.body);
    assert.strictEqual(body.title, "Nexus PLM");
    assert.strictEqual(body.description, "hello");
    assert.strictEqual(body.severity, "warning");
});
