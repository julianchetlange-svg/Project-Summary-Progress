/**
 * McCoy's Project Summary Report — Gmail ↔ GitHub handoff
 * Install in script.google.com while signed in as mccoysfleet@gmail.com.
 *
 * Every 5 minutes:
 *   1. New Fexa report emails  → attachment uploaded to the private inbox repo,
 *      then the report workflow is started (repository_dispatch).
 *   2. Emails the workflow left in the inbox repo's outbox/ → sent from this
 *      Gmail account, then deleted from outbox/.
 *
 * Setup: Project Settings → Script properties → add GITHUB_TOKEN (fine-grained
 * token with Contents: Read and write on both repos). Then run setup() once.
 */
const CFG = {
  owner: 'julianchetlange-svg',
  inboxRepo: 'Project-Summary-Inbox',      // private: raw exports + outgoing emails
  reportRepo: 'Project-Summary-Progress',  // public: the report + workflows
  sender: 'donotreply@fexa.io',
  kinds: {
    module: 'Project Report - Module',
    contingency: 'Current Budget Year Contingency',
  },
  label: 'Fexa/Uploaded',
  tz: 'America/Chicago',
  fromName: "McCoy's Project Report",
};

/** Run once by hand: authorizes, records the start time and creates the 5-minute trigger. */
function setup() {
  if (!token_()) throw new Error('Add the GITHUB_TOKEN script property first (Project Settings → Script properties).');
  gh_('GET', `/repos/${CFG.owner}/${CFG.inboxRepo}`);   // fails loudly if the token can't see the inbox repo
  gh_('GET', `/repos/${CFG.owner}/${CFG.reportRepo}`);
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('INSTALLED_AT')) props.setProperty('INSTALLED_AT', String(Date.now()));
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('tick').timeBased().everyMinutes(5).create();
  Logger.log('Set up. Emails received after %s will be handed off.', new Date(Number(props.getProperty('INSTALLED_AT'))));
}

/** Main loop (runs on the trigger). */
function tick() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return;
  try {
    uploadNewReports_();
    sendOutbox_();
  } finally {
    lock.releaseLock();
  }
}

function uploadNewReports_() {
  const props = PropertiesService.getScriptProperties();
  const since = Number(props.getProperty('INSTALLED_AT') || Date.now());
  const label = GmailApp.getUserLabelByName(CFG.label) || GmailApp.createLabel(CFG.label);
  for (const kind of Object.keys(CFG.kinds)) {
    const q = `from:${CFG.sender} subject:"${CFG.kinds[kind]}" has:attachment newer_than:3d`;
    for (const thread of GmailApp.search(q, 0, 20)) {
      for (const m of thread.getMessages()) {
        const key = 'done_' + m.getId();
        if (props.getProperty(key) || m.getDate().getTime() < since) continue;
        if (m.getFrom().indexOf(CFG.sender) < 0) continue;
        const day = Utilities.formatDate(m.getDate(), CFG.tz, 'yyyy-MM-dd');
        const hms = Utilities.formatDate(m.getDate(), CFG.tz, 'HHmmss');
        let uploaded = 0;
        for (const a of m.getAttachments()) {
          const ext = (a.getName().match(/\.(xlsx?|csv)$/i) || [])[0];
          if (!ext) continue;
          const path = `incoming/${kind}/${day}_${hms}_${m.getId()}${ext.toLowerCase()}`;
          if (!gh_('GET', `/repos/${CFG.owner}/${CFG.inboxRepo}/contents/${path}`)) {
            gh_('PUT', `/repos/${CFG.owner}/${CFG.inboxRepo}/contents/${path}`, {
              message: `Fexa ${kind} export ${day} ${hms}`,
              content: Utilities.base64Encode(a.getBytes()),
            });
          }
          uploaded++;
        }
        if (uploaded) {
          gh_('POST', `/repos/${CFG.owner}/${CFG.reportRepo}/dispatches`, {
            event_type: `fexa-${kind}`,
            client_payload: { date: day },
          });
          thread.addLabel(label);
          Logger.log('Uploaded %s export from %s', kind, day);
        }
        props.setProperty(key, String(Date.now()));
      }
    }
  }
}

function sendOutbox_() {
  const listing = gh_('GET', `/repos/${CFG.owner}/${CFG.inboxRepo}/contents/outbox`);
  if (!listing) return;
  for (const f of listing.filter(x => x.type === 'file' && /\.json$/.test(x.name))) {
    const file = gh_('GET', `/repos/${CFG.owner}/${CFG.inboxRepo}/contents/${f.path}`);
    const msg = JSON.parse(Utilities.newBlob(Utilities.base64Decode(file.content.replace(/\n/g, ''))).getDataAsString('UTF-8'));
    GmailApp.sendEmail(msg.to.join(','), msg.subject, 'This email is best viewed in an HTML-capable client.', {
      htmlBody: msg.html,
      name: CFG.fromName,
    });
    gh_('DELETE', `/repos/${CFG.owner}/${CFG.inboxRepo}/contents/${f.path}`, {
      message: `Sent: ${msg.subject}`,
      sha: file.sha,
    });
    Logger.log('Sent %s', msg.subject);
  }
}

function token_() {
  return PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
}

/** GitHub REST call. Returns parsed JSON, {} for empty bodies, or null for 404. */
function gh_(method, path, body) {
  const res = UrlFetchApp.fetch('https://api.github.com' + path, {
    method: method.toLowerCase(),
    headers: {
      Authorization: 'Bearer ' + token_(),
      Accept: 'application/vnd.github+json',
      'X-GitHub-Api-Version': '2022-11-28',
    },
    contentType: 'application/json',
    payload: body ? JSON.stringify(body) : undefined,
    muteHttpExceptions: true,
  });
  const code = res.getResponseCode();
  if (code === 404) return null;
  if (code >= 300) throw new Error(`${method} ${path} → ${code}: ${res.getContentText().slice(0, 300)}`);
  const text = res.getContentText();
  return text ? JSON.parse(text) : {};
}
