// resellpro-valor-webhook — instant ACK + background forward
export default {
  async fetch(request, env, ctx) {
    // Small JSON 200 used by Valor's validator (GET/HEAD/OPTIONS)
    const ack = () =>
      new Response(
        JSON.stringify({
          ok: true,
          service: "Resell Pro",
          hook: "valor",
          ts: new Date().toISOString(),
        }),
        { status: 200, headers: { "content-type": "application/json" } }
      );

    // Many validators probe with HEAD/OPTIONS/GET
    if (request.method === "HEAD" || request.method === "OPTIONS" || request.method === "GET") {
      return ack();
    }

    if (request.method === "POST") {
      const body = await request.text();

      // Forward to Apps Script **in the background** (do NOT await)
      // Add ?source=valor so your doPost routes it to valorWebhookHandler_
      const target =
      //"https://script.google.com/macros/s/AKfycbwCIPpnLNDYX0xQgiSfULrbfcDz38MXUh-dViFfY1E/dev?source=valor"
      "https://script.google.com/macros/s/AKfycbwx5QMnLUWpbQJBe4FczbEJEZJxxrcHWBqmdOL09GtH9_x5B03tgplFuSQg8ViG2Etw/exec?source=valor";
      
      ctx.waitUntil(forwardWithRetry(target, body));

      // Immediate 200 so Valor’s webhook validation never waits >2s
      return ack();
    }

    // Harmless default
    return ack();
  },
};

// Fire-and-forget with a few retries in case Apps Script is cold/slow
async function forwardWithRetry(url, body) {
  const delays = [0, 500, 1500, 5000]; // ms
  for (const ms of delays) {
    if (ms) await sleep(ms);
    try {
      const r = await fetch(url, {
        method: "POST",
        headers: { "content-type": "application/json" },
        body,
      });
      if (r.ok) return; // delivered
    } catch (_) {
      // ignore and retry
    }
  }
  // (Optional) add KV/Queues here if you want guaranteed persistence
}
function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }
