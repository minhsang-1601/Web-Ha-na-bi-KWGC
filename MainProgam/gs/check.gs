function whoAmI() {
  console.log(Session.getEffectiveUser().getEmail());
  console.log(MailApp.getRemainingDailyQuota());
}
function checkQuotaType() {
  const email = Session.getEffectiveUser().getEmail();
  const quota = MailApp.getRemainingDailyQuota();
  console.log('Email:', email);
  console.log('Quota 残数:', quota);
}
function checkSentToday() {
  const threads = GmailApp.search('from:s-tm-tran@itsac.biz after:2026/06/22');
  console.log('本日送信済みメール件数:', threads.length);
}
function checkWorkspaceInfo() {
  console.log('Email:', Session.getEffectiveUser().getEmail());
  console.log('Quota:', MailApp.getRemainingDailyQuota());
  const domain = Session.getEffectiveUser().getEmail().split('@')[1];
  console.log('Domain:', domain);
}