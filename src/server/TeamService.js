/**
 * TeamService.js — the head's and the auditor's view of the whole department
 * (SPEC §8.4): who is carrying what, what has slipped, and what already has its
 * PR number and is only waiting to be closed.
 */
var TeamService = (function () {

  function overview(user) {
    Auth.requireRole(user, [Auth.ROLES.HEAD, Auth.ROLES.AUDITOR, Auth.ROLES.ADMIN]);

    var openCases = Repository.query('Cases', {
      where: function (c) { return !Auth.isTerminal(c.Status); }
    });
    var nextActions = CaseService.nextActionByCase(openCases.map(function (c) { return c.Case_ID; }));
    var readiness = Rules.summarizeCases(openCases);

    var buyers = {};
    Auth.listBuyers().forEach(function (b) {
      buyers[b.email] = {
        email: b.email, name: b.name, role: b.role,
        openCases: 0, overdueActions: 0, readyToClose: 0, cases: []
      };
    });

    openCases.forEach(function (c) {
      var owner = String(c.Buyer_Owner).trim().toLowerCase();
      if (!buyers[owner]) {
        // A Case whose owner has since left or been deactivated still has to be visible.
        buyers[owner] = {
          email: owner, name: owner, role: '—',
          openCases: 0, overdueActions: 0, readyToClose: 0, cases: [], inactive: true
        };
      }
      var bucket = buyers[owner];
      var next = nextActions[c.Case_ID] || null;
      var ready = readiness[c.Case_ID] || null;

      bucket.openCases++;
      if (next && next.overdue) bucket.overdueActions++;
      if (ready && ready.ok) bucket.readyToClose++;

      bucket.cases.push(Object.assign(Repository.toClient(c), { nextAction: next, readiness: ready }));
    });

    var rows = Object.keys(buyers).map(function (email) { return buyers[email]; });
    rows.sort(function (a, b) { return b.openCases - a.openCases; });
    rows.forEach(function (row) {
      row.cases.sort(function (a, b) { return String(a.Case_ID).localeCompare(String(b.Case_ID)); });
    });

    return {
      buyers: rows,
      totals: {
        openCases: openCases.length,
        overdueActions: rows.reduce(function (n, r) { return n + r.overdueActions; }, 0),
        readyToClose: rows.reduce(function (n, r) { return n + r.readyToClose; }, 0)
      },
      canReassign: Auth.isHead(user)
    };
  }

  return { overview: overview };
})();
