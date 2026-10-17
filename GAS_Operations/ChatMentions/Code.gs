/**
 * GCX Chat mentions — builds the email → Google Chat user ID map ("users/123…") that real @mentions need.
 *
 * Google Chat only pings someone when a message's plain text contains <users/USER_ID> (like <users/all>).
 * Run buildMentionMap() once (and again when staff change). It:
 *   1. lists the members of every room in GCX_SPACES (Chat advanced service, runs as you), and
 *   2. resolves each member's email through the Workspace directory (People advanced service),
 *      plus any extra emails in EXTRA_EMAILS that aren't in those rooms yet;
 * then saves {email: "users/ID"} to script property MENTION_MAP and to Drive as gcx_chat_user_ids.json
 * (copy that file to ~/.config/ on the Macs for the Python senders — ClaudeConfig/lib/chat_mentions.py).
 *
 * Script properties:
 *   GCX_SPACES    comma-separated room ids, e.g. "spaces/AAQAdqYt1ro,spaces/AAQAc9NQmJQ"
 *   EXTRA_EMAILS  optional comma-separated staff emails
 * Advanced services (appsscript.json): Chat v1, People v1.
 */

function buildMentionMap() {
  const props = PropertiesService.getScriptProperties();
  const spaces = (props.getProperty('GCX_SPACES') || '').split(',').map(s => s.trim()).filter(String);
  const map = {};
  const ids = new Set();

  // 1. room members → user ids (membership lists give ids; emails come from the directory)
  spaces.forEach(space => {
    let pageToken;
    do {
      const res = Chat.Spaces.Members.list(space, { pageSize: 100, pageToken: pageToken, filter: 'member.type = "HUMAN"' });
      (res.memberships || []).forEach(m => { if (m.member && m.member.name) ids.add(m.member.name); });
      pageToken = res.nextPageToken;
    } while (pageToken);
  });
  ids.forEach(name => {
    const email = emailForUser_(name);
    if (email) map[email] = name;
  });

  // 2. extra staff by email → directory lookup
  (props.getProperty('EXTRA_EMAILS') || '').split(',').map(s => s.trim().toLowerCase()).filter(String).forEach(email => {
    if (!map[email]) { const name = userForEmail_(email); if (name) map[email] = name; }
  });

  props.setProperty('MENTION_MAP', JSON.stringify(map));
  const blob = Utilities.newBlob(JSON.stringify(map, null, 1), 'application/json', 'gcx_chat_user_ids.json');
  const existing = DriveApp.getFilesByName('gcx_chat_user_ids.json');
  if (existing.hasNext()) existing.next().setContent(blob.getDataAsString()); else DriveApp.createFile(blob);
  Logger.log('mention map: %s people from %s rooms', Object.keys(map).length, spaces.length);
  return map;
}

/** "<users/123>" for an email (from the saved map, else a live directory lookup), "@name" when unknown. */
function mentionByEmail(email) {
  if (String(email).toLowerCase() === 'all') return '<users/all>';
  const map = JSON.parse(PropertiesService.getScriptProperties().getProperty('MENTION_MAP') || '{}');
  const key = String(email).trim().toLowerCase();
  const name = map[key] || userForEmail_(key);
  return name ? '<' + name + '>' : '@' + key.split('@')[0];
}

/** Prepends mentions to a webhook/app message payload's plain text (cards stay below it). */
function withMentions(payload, emails) {
  const tags = (emails || []).map(mentionByEmail).join(' ');
  payload.text = [tags, payload.text || ''].filter(String).join(' ').trim();
  return payload;
}

// users/123 → primary email, via the directory (people/123 is the same id)
function emailForUser_(userName) {
  try {
    const p = People.People.get('people/' + userName.split('/')[1], { personFields: 'emailAddresses' });
    const e = (p.emailAddresses || []).find(x => x.metadata && x.metadata.primary) || (p.emailAddresses || [])[0];
    return e ? e.value.toLowerCase() : null;
  } catch (err) { Logger.log('no email for %s: %s', userName, err); return null; }
}

// email → users/123, via the domain directory
function userForEmail_(email) {
  try {
    const res = People.People.searchDirectoryPeople({ query: email, readMask: 'emailAddresses', sources: ['DIRECTORY_SOURCE_TYPE_DOMAIN_PROFILE'] });
    const hit = (res.people || []).find(p => (p.emailAddresses || []).some(x => x.value.toLowerCase() === email));
    return hit ? 'users/' + hit.resourceName.split('/')[1] : null;
  } catch (err) { Logger.log('directory lookup failed for %s: %s', email, err); return null; }
}

/** Test: posts one mention to a room via its webhook URL (script property TEST_WEBHOOK). */
function testMention() {
  const hook = PropertiesService.getScriptProperties().getProperty('TEST_WEBHOOK');
  const me = Session.getActiveUser().getEmail();
  const payload = withMentions({ text: '멘션 테스트입니다 (GCX ChatMentions).' }, [me]);
  const res = UrlFetchApp.fetch(hook, { method: 'post', contentType: 'application/json; charset=UTF-8', payload: JSON.stringify(payload), muteHttpExceptions: true });
  Logger.log('%s %s', res.getResponseCode(), res.getContentText().slice(0, 300));
}
