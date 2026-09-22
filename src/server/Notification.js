/**
 * Notification.js — e-mail, and the daily job behind it (SPEC §11).
 *
 * Two rules shape everything here:
 *   PDPA — a notification never carries vendor contact details. It carries the
 *   Case, what happened, and a link; whoever needs the contact opens the app.
 *   A failed send must never fail the write that triggered it, so every send is
 *   wrapped and logged rather than thrown.
 */
var Notification = (function () {

  var FROM_NAME = 'ระบบติดตามกิจกรรมจัดซื้อ';

  /* --------------------------------------------------------------- sending */

  function send(to, subject, lines) {
    var recipients = (Array.isArray(to) ? to : [to])
      .map(function (e) { return String(e || '').trim(); })
      .filter(function (e) { return !!e; });
    if (recipients.length === 0) return false;

    try {
      MailApp.sendEmail({
        to: Utils.unique(recipients).join(','),
        subject: subject,
        name: FROM_NAME,
        htmlBody: htmlBody(lines),
        body: lines.map(stripTags).join('\n')
      });
      return true;
    } catch (e) {
      console.error('sendEmail to ' + recipients.join(',') + ' failed: ' + ((e && e.message) || e));
      return false;
    }
  }

  function htmlBody(lines) {
    return '<div style="font-family:Sarabun,\'Noto Sans Thai\',sans-serif;font-size:15px;line-height:1.7">' +
      lines.join('') +
      '<p style="color:#888;font-size:12px;margin-top:24px">' +
      'อีเมลฉบับนี้ส่งโดยอัตโนมัติจาก' + FROM_NAME + ' กรุณาอย่าตอบกลับ</p></div>';
  }

  function stripTags(html) {
    return String(html).replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ').trim();
  }

  function appUrl() {
    try {
      return ScriptApp.getService().getUrl() || '';
    } catch (e) {
      return '';
    }
  }

  function caseLink(caseRecord) {
    var url = appUrl();
    var label = caseRecord.Case_ID + ' — ' + Utils.truncate(caseRecord.Description, 80);
    if (!url) return '<p><strong>' + escape(label) + '</strong></p>';
    return '<p><a href="' + url + '#/case/' + encodeURIComponent(caseRecord.Case_ID) + '">' +
      escape(label) + '</a></p>';
  }

  /** Mail is HTML, so anything that came from a user gets escaped. */
  function escape(value) {
    return String(value === null || value === undefined ? '' : value)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  function headEmails() {
    return Auth.listActiveUsers()
      .filter(function (u) { return u.role === Auth.ROLES.HEAD; })
      .map(function (u) { return u.email; });
  }

  /* -------------------------------------------------------------- triggers */

  /** The Case owner hears about a status change somebody else made. */
  function onStatusChanged(caseRecord, from, to, user) {
    if (String(caseRecord.Buyer_Owner).toLowerCase() === user.email) return;
    send(caseRecord.Buyer_Owner, '[' + caseRecord.Case_ID + '] สถานะงานถูกเปลี่ยนเป็น ' + Auth.statusLabel(to),
      [
        '<p>' + escape(user.name) + ' เปลี่ยนสถานะงานที่คุณเป็นเจ้าของ</p>',
        caseLink(caseRecord),
        '<p>จาก <strong>' + escape(Auth.statusLabel(from)) + '</strong> เป็น <strong>' +
          escape(Auth.statusLabel(to)) + '</strong></p>'
      ]);
  }

  /** SPEC §6.1 — the owner must know the system pulled their Case back. */
  function onAutoRevert(caseRecord, from, to, message) {
    send(caseRecord.Buyer_Owner, '[' + caseRecord.Case_ID + '] ระบบย้อนสถานะงานอัตโนมัติ', [
      '<p>ระบบย้อนสถานะงานของคุณกลับจาก <strong>' + escape(Auth.statusLabel(from)) +
        '</strong> เป็น <strong>' + escape(Auth.statusLabel(to)) + '</strong></p>',
      caseLink(caseRecord),
      '<p>เหตุผล: ' + escape(message) + '</p>',
      '<p>กรุณาตรวจสอบและกรอกข้อมูลให้ครบก่อนปิดงานอีกครั้ง</p>'
    ]);
  }

  /** Acceptance Test 15 — the new owner is told, and the previous one too. */
  function onReassigned(caseRecord, previousOwner, newOwner, user, reason) {
    send(newOwner, '[' + caseRecord.Case_ID + '] คุณได้รับมอบหมายงานจัดซื้อใหม่', [
      '<p>' + escape(user.name) + ' โอนงานนี้มาให้คุณดูแล</p>',
      caseLink(caseRecord),
      '<p>เจ้าของงานเดิม: ' + escape(previousOwner) + '</p>',
      '<p>เหตุผล: ' + escape(reason) + '</p>'
    ]);
    send(previousOwner, '[' + caseRecord.Case_ID + '] งานถูกโอนให้ผู้อื่นดูแลแล้ว', [
      '<p>' + escape(user.name) + ' โอนงานนี้ไปให้ ' + escape(newOwner) + '</p>',
      caseLink(caseRecord),
      '<p>เหตุผล: ' + escape(reason) + '</p>'
    ]);
  }

  function onReopened(caseRecord, from, to, user, reason) {
    send(caseRecord.Buyer_Owner, '[' + caseRecord.Case_ID + '] งานถูกเปิดขึ้นมาใหม่', [
      '<p>' + escape(user.name) + ' เปิดงานที่ปิดไปแล้วขึ้นมาใหม่</p>',
      caseLink(caseRecord),
      '<p>กลับไปที่สถานะ <strong>' + escape(Auth.statusLabel(to)) + '</strong></p>',
      '<p>เหตุผล: ' + escape(reason) + '</p>'
    ]);
  }

  /* ------------------------------------------------------------ daily job */

  /**
   * One digest per buyer: next actions already overdue, and those falling due
   * within REMINDER_DAYS_AHEAD (SPEC §11).
   */
  function buildDailyDigests() {
    var daysAhead = Config.getNumber('REMINDER_DAYS_AHEAD', 1);
    var today = Utils.today();
    var horizon = Utils.addDays(today, daysAhead);

    var cases = {};
    Repository.query('Cases', {
      where: function (c) { return !Auth.isTerminal(c.Status); }
    }).forEach(function (c) { cases[c.Case_ID] = c; });

    var digests = {};
    ActivityService.outstandingNextActions(function (a) {
      return Object.prototype.hasOwnProperty.call(cases, a.Case_ID);
    }).forEach(function (activity) {
      var due = Utils.startOfDay(activity.Next_Action_Date);
      if (!due || due.getTime() > horizon.getTime()) return;

      var caseRecord = cases[activity.Case_ID];
      var owner = String(caseRecord.Buyer_Owner).trim().toLowerCase();
      if (!digests[owner]) digests[owner] = { overdue: [], upcoming: [] };

      var entry = {
        caseId: caseRecord.Case_ID,
        description: caseRecord.Description,
        text: activity.Next_Action,
        due: due,
        caseRecord: caseRecord
      };
      if (due.getTime() < today.getTime()) digests[owner].overdue.push(entry);
      else digests[owner].upcoming.push(entry);
    });

    Object.keys(digests).forEach(function (owner) {
      digests[owner].overdue.sort(byDue);
      digests[owner].upcoming.sort(byDue);
    });
    return digests;
  }

  function byDue(a, b) {
    return a.due.getTime() - b.due.getTime();
  }

  function sendDailyDigests() {
    var digests = buildDailyDigests();
    var sent = 0;
    Object.keys(digests).forEach(function (owner) {
      var digest = digests[owner];
      var total = digest.overdue.length + digest.upcoming.length;
      if (total === 0) return;

      var lines = ['<p>สรุปสิ่งที่ต้องทำต่อในงานจัดซื้อของคุณ</p>'];
      if (digest.overdue.length) {
        lines.push('<h3 style="color:#b3261e">เลยกำหนดแล้ว (' + digest.overdue.length + ')</h3>');
        lines.push(listHtml(digest.overdue));
      }
      if (digest.upcoming.length) {
        lines.push('<h3>ครบกำหนดเร็วๆ นี้ (' + digest.upcoming.length + ')</h3>');
        lines.push(listHtml(digest.upcoming));
      }
      if (send(owner, 'สรุปงานจัดซื้อที่ต้องติดตาม ' + Utils.formatDate(Utils.today()), lines)) sent++;
    });
    return { recipients: Object.keys(digests).length, sent: sent };
  }

  function listHtml(entries) {
    return '<ul>' + entries.map(function (e) {
      return '<li><strong>' + escape(e.caseId) + '</strong> — ' + escape(e.text) +
        ' (ภายใน ' + escape(Utils.formatDate(e.due)) + ')<br>' +
        '<span style="color:#666">' + escape(Utils.truncate(e.description, 80)) + '</span></li>';
    }).join('') + '</ul>';
  }

  /**
   * A closed Case is rechecked once a day as well as after every write, so that
   * an edit made straight in the sheet — the one path that bypasses the rules —
   * does not leave a Case sitting closed against a PR number it no longer has.
   */
  function recheckOpenCases() {
    var reverted = 0;
    Repository.query('Cases', {
      where: function (c) { return c.Status === 'CLOSED'; }
    }).forEach(function (caseRecord) {
      if (StatusEngine.runRechecks(caseRecord).reverted) reverted++;
    });
    return reverted;
  }

  return {
    send: send,
    appUrl: appUrl,
    headEmails: headEmails,
    onStatusChanged: onStatusChanged,
    onAutoRevert: onAutoRevert,
    onReassigned: onReassigned,
    onReopened: onReopened,
    buildDailyDigests: buildDailyDigests,
    sendDailyDigests: sendDailyDigests,
    recheckOpenCases: recheckOpenCases
  };
})();

/**
 * The time-driven trigger installed by installTriggers(). Kept as a top-level
 * function because Apps Script triggers can only call one.
 */
function dailyReminderJob() {
  var reverted = Notification.recheckOpenCases();
  var digests = Notification.sendDailyDigests();
  console.log('dailyReminderJob: ' + reverted + ' cases auto-reverted, ' +
    digests.sent + '/' + digests.recipients + ' reminder e-mails sent');
  return { reverted: reverted, digests: digests };
}
