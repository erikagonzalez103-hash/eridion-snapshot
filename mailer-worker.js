// ════════════════════════════════════════════════════════════════════════
//  ERIDION SNAPSHOT MAILER — Cloudflare Worker
//  Emails Dion the weekly snapshot every Monday ~6am Central.
//
//  How it works: reads the headline numbers from the live snapshot's
//  data.json (kept fresh by Claude), renders an email, and sends it via
//  Microsoft Graph using your existing Eridion 365 app registration.
//  No Xero connection lives in this worker — it only needs Microsoft creds.
//
//  ── ONE-TIME DEPLOY (all in the Cloudflare dashboard, no wrangler) ──
//  1. Workers & Pages → Create → Worker → name it "eridion-snapshot-mailer"
//     → Deploy → Edit code → paste this whole file → Deploy.
//  2. The worker's Settings → Variables and Secrets → add:
//       MS_TENANT_ID       (Text)    your Microsoft directory (tenant) ID
//       MS_CLIENT_ID       (Text)    the Azure app (client) ID
//       MS_CLIENT_SECRET   (Secret)  the Azure app client secret*
//       SENDER_EMAIL       (Text)    erika@eridionglass.com
//       DION_EMAIL         (Text)    dion@eridionglass.com
//       SNAPSHOT_JSON_URL  (Text)    https://erikagonzalez103-hash.github.io/eridion-snapshot/data.json
//     *If you don't have the secret value, make a new one in Azure:
//      App registrations → (your app) → Certificates & secrets → New client secret.
//     The app needs the "Mail.Send" Application permission (it already has it —
//     that's how the old Francisco worker sent mail).
//  3. Settings → Triggers → Cron Triggers → Add Cron Trigger:  0 11 * * 1
//       (Monday 11:00 UTC = 6:00am CDT / 5:00am CST — safely before Dion leaves.)
//
//  ── TEST BEFORE RELYING ON IT ──
//    https://<your-worker>.workers.dev/preview    → shows the email, sends nothing
//    https://<your-worker>.workers.dev/send-now   → sends one to Dion right now
// ════════════════════════════════════════════════════════════════════════

const GRAPH = 'https://graph.microsoft.com/v1.0';
const money = n => '$' + Math.round(n).toLocaleString('en-US');

async function getMSToken(env) {
  const r = await fetch('https://login.microsoftonline.com/' + env.MS_TENANT_ID + '/oauth2/v2.0/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      client_id: env.MS_CLIENT_ID,
      client_secret: env.MS_CLIENT_SECRET,
      scope: 'https://graph.microsoft.com/.default',
      grant_type: 'client_credentials',
    }),
  });
  const d = await r.json();
  if (!d.access_token) throw new Error('Microsoft token failed: ' + JSON.stringify(d));
  return d.access_token;
}

async function loadData(env) {
  // cache-bust so we always get the freshest data.json
  const r = await fetch(env.SNAPSHOT_JSON_URL + '?t=' + Date.now(), { cf: { cacheTtl: 0 } });
  if (!r.ok) throw new Error('Could not load data.json: HTTP ' + r.status);
  return r.json();
}

function buildEmail(d) {
  const behind = d.pacePct < 0;
  return `<div style="font-family:-apple-system,Segoe UI,sans-serif;color:#1a1a1a;max-width:560px">
  <p>Good morning Dion,</p>
  <p>Where Eridion stands this week — <strong>cash basis</strong> (money actually collected):</p>
  <table style="border-collapse:collapse;font-size:14px;margin:12px 0">
    <tr><td style="padding:6px 14px 6px 0"><strong>${d.lastMonthName} closed</strong></td>
        <td style="padding:6px 0;color:#0F6E56;font-weight:600">${money(d.lastMonthCollected)}</td></tr>
    <tr><td style="padding:6px 14px 6px 0"><strong>2026 YTD</strong></td>
        <td style="padding:6px 0;font-weight:600">${money(d.ytd)}</td>
        <td style="padding:6px 0 6px 12px;color:${behind ? '#A32D2D' : '#0F6E56'}">${d.pacePct >= 0 ? '+' : ''}${d.pacePct}% vs 2025 (${money(d.ytdPriorYear)} by now)</td></tr>
    <tr><td style="padding:6px 14px 6px 0">Goal forecast</td>
        <td style="padding:6px 0">${money(d.goalForecast)}</td>
        <td style="padding:6px 0 6px 12px;color:#666">+5%/mo push · 2025 finished ${money(d.priorYearTotal)}</td></tr>
  </table>
  <p style="font-size:13px;color:#666">Owed to us (A/R): ~${money(d.arOwed)} <em>(inflated by ghost invoices — real number is lower)</em> &nbsp;·&nbsp; Bills we owe: ${money(d.apOwed)}</p>
  <p><a href="${d.url}" style="background:#C8963E;color:#0D1B2A;padding:9px 18px;border-radius:6px;text-decoration:none;font-weight:600;font-size:13px">Open full snapshot →</a></p>
  <p style="color:#888;font-size:13px">— Eridion Snapshot · sent automatically Monday morning</p>
</div>`;
}

async function sendToDion(env, html) {
  const token = await getMSToken(env);
  const r = await fetch(GRAPH + '/users/' + env.SENDER_EMAIL + '/sendMail', {
    method: 'POST',
    headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' },
    body: JSON.stringify({
      message: {
        subject: 'Eridion — Monday Snapshot',
        body: { contentType: 'HTML', content: html },
        toRecipients: [{ emailAddress: { address: env.DION_EMAIL } }],
      },
      saveToSentItems: true,
    }),
  });
  if (!r.ok) throw new Error('sendMail failed: HTTP ' + r.status + ' — ' + (await r.text()));
}

export default {
  // Fires on the cron schedule (Monday 11:00 UTC).
  async scheduled(event, env, ctx) {
    ctx.waitUntil((async () => {
      const d = await loadData(env);
      await sendToDion(env, buildEmail(d));
    })());
  },
  // Manual endpoints for testing.
  async fetch(req, env) {
    const path = new URL(req.url).pathname;
    try {
      if (path === '/preview') {
        return new Response(buildEmail(await loadData(env)), { headers: { 'Content-Type': 'text/html; charset=utf-8' } });
      }
      if (path === '/send-now') {
        await sendToDion(env, buildEmail(await loadData(env)));
        return new Response('Sent to ' + env.DION_EMAIL + ' ✓');
      }
      return new Response('Eridion Snapshot Mailer is live. Try /preview (no send) or /send-now (sends to Dion). Cron: Mondays 11:00 UTC.');
    } catch (e) {
      return new Response('Error: ' + e.message, { status: 500 });
    }
  },
};
