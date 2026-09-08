/*
 * Event-based runtime: runs on every new message compose and writes the
 * signature, with the out-of-office block on top.
 *
 * Hard rule for this file: NO network calls and NO sign-in prompts. Outlook
 * gives a compose handler only a few seconds before it kills the runtime, and a
 * blocked handler means the user gets no signature at all. Everything needed is
 * read from roaming settings, which the task pane refreshes.
 *
 * Dates are stored rather than rendered HTML, so days that have passed since the
 * last refresh are dropped here, offline, at the moment the mail is composed.
 */

var STORE = {
  days: 'oooDays',
  signature: 'signatureHtml',
  options: 'oooOptions',
  enabled: 'oooEnabled'
};

function readSettings() {
  var rs = Office.context.roamingSettings;
  var days = {}, options = {};
  try { days = JSON.parse(rs.get(STORE.days) || '{}'); } catch (e) { days = {}; }
  try { options = JSON.parse(rs.get(STORE.options) || '{}'); } catch (e) { options = {}; }
  return {
    days: days,
    options: options,
    signature: rs.get(STORE.signature) || '',
    enabled: rs.get(STORE.enabled) !== false
  };
}

function buildSignatureHtml() {
  var s = readSettings();
  if (!s.signature && !s.enabled) { return null; }

  var block = '';
  if (s.enabled) {
    try {
      block = OooCore.blockFromDays(s.days, s.options) || '';
    } catch (e) {
      block = ''; // never let a rendering fault cost the user their signature
    }
  }
  var html = block ? (block + s.signature) : s.signature;
  return html || null;
}

function onNewMessageComposeHandler(event) {
  var done = false;
  function finish() {
    if (done) { return; }
    done = true;
    event.completed();
  }

  // Belt and braces: if anything stalls, still release the compose window.
  setTimeout(finish, 4000);

  try {
    var html = buildSignatureHtml();
    if (!html) { finish(); return; }

    Office.context.mailbox.item.body.setSignatureAsync(
      html,
      { coercionType: Office.CoercionType.Html },
      function () { finish(); }
    );
  } catch (e) {
    finish();
  }
}

// Outlook on Windows (classic) uses this association; OWA and new Outlook pick
// the function up from the manifest's FunctionName.
if (typeof Office !== 'undefined' && Office.actions && Office.actions.associate) {
  Office.actions.associate('onNewMessageComposeHandler', onNewMessageComposeHandler);
}
