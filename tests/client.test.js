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
    // Not "signed out" and not "not in PLM": with no secret the service refuses every call, so
    // either of those would be a guess about a service that never answered. A bare 403 would send
    // somebody looking at permissions in PLM instead of at the install step.
    assert.strictEqual(
        panel.headline({ status: "checked_in", part_number: "NX1" }, { hasSecret: false }),
        panel.NOT_INSTALLED
    );
});

test("the missing secret is reported before anything else", () => {
    assert.strictEqual(
        panel.headline(null, { hasSecret: false, signedIn: false, hasDocument: false }),
        panel.NOT_INSTALLED
    );
});

test("with a secret the panel goes back to talking about the document", () => {
    assert.strictEqual(
        panel.headline({ status: "checked_in", part_number: "NX1" }, { hasSecret: true }),
        "NX1"
    );
});

test("a service that is not answering is not reported as signed out", () => {
    // Each of these is fixed by a different action. Saying "sign in" to all three sent people to
    // a dialog that could not help.
    assert.strictEqual(panel.headline(null, { unreachable: true }), panel.NOT_RUNNING);
    assert.strictEqual(panel.headline(null, { refused: true }), panel.REFUSED);
    assert.strictEqual(panel.headline(null, { signedIn: false }), panel.NOT_SIGNED_IN);
    assert.strictEqual(panel.headline(null, { hasSecret: false }), panel.NOT_INSTALLED);
});

test("the status travels with the answer so the panel can tell them apart", async () => {
    const client = new clientModule.Client(undefined,
        () => Promise.resolve({ status: 403, text: () => Promise.resolve('{"success":false}') }));

    const answer = await client.me();

    assert.strictEqual(answer.httpStatus, 403);
    assert.strictEqual(answer.refused, true);
});

test("the HTTP status never overwrites the item's own status", () => {
    // `status` on a /plm/state answer is the lifecycle. Putting the HTTP status there showed
    // "400" in the panel's Status row and made a document PLM has never seen look like one it
    // knew, because isInPlm only asks whether status is set and not "unknown".
    const answer = { success: false, error: "file_path or item_id is required" };
    answer.httpStatus = 400;

    assert.strictEqual(panel.isInPlm(answer), false);
    assert.strictEqual(panel.rowsFor(answer).find(([l]) => l === "Status")[1], panel.ABSENT);
});

// ── the tray ─────────────────────────────────────────────────────────────────

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
