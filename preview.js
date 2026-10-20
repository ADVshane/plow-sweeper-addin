/* Preview mode: runs only when the page is opened directly (e.g. on GitHub Pages), never inside MyGeotab.
   Supplies a fake API with sample data so the UI can be reviewed in a browser. Nothing is saved. */
(function () {
  'use strict';
  if (/(^|\.)geotab\.com$/i.test(location.hostname)) return;   // inside MyGeotab: do nothing

  function start() {
    if (window.__psiLive) return;                                 // real MyGeotab API already initialized
    const now = Date.now(), day = 864e5, iso = t => new Date(t).toISOString();
    const runs = [
      { runName: 'Main St plow check', presets: ['plow'], aux: [], req: 'r1', from: now - 14 * day, created: now - day },
      { runName: 'Downtown sweeper route', presets: ['sweeper'], aux: [], req: 'r2', from: now - 30 * day, created: now - 2 * 3600e3 },
      { runName: 'Industrial park spreader', presets: ['spreader'], aux: [5], req: 'r3', from: now - 5 * day, created: now - 3 * day }
    ];
    const requests = { r1: 'Complete', r2: 'Scheduled', r3: 'Cancelled' };
    const children = {
      r1: Array(42).fill({ status: 'Success' }),
      r2: Array(17).fill({ status: 'Success' }).concat(Array(25).fill({ status: 'Scheduled' })),
      r3: Array(42).fill({ status: 'Cancelled' })
    };
    const trucks = [['b11', 'Plow 11'], ['b12', 'Plow 12'], ['b17', 'Plow 17'], ['b23', 'Tandem 23']];
    const events = [];
    for (let i = 0; i < 9; i++) {
      const t = now - (13 - i) * day + (6 + i % 5) * 3600e3, d = trucks[i % trucks.length][0];
      events.push({ device: { id: d }, activeFrom: iso(t), activeTo: iso(t + (12 + i * 7) * 60e3) });
    }

    function get(p) {
      const s = p.search || {};
      switch (p.typeName) {
        case 'AddInData':
          return runs.map((r, i) => ({ id: 'd' + i, details: {
            runName: r.runName, presets: r.presets, aux: r.aux, reprocessRequestId: r.req, ruleId: 'rule' + i, zoneId: 'z' + i,
            fromDate: iso(r.from), toDate: iso(now), createdAt: iso(r.created), createdBy: 'shane@advtracking.net', cleanedUp: false } }));
        case 'ReprocessRequest': return [{ id: s.id, status: requests[s.id] }];
        case 'DeviceReprocessRequest': return children[s.reprocessRequestId] || [];
        case 'ExceptionEvent': return events;
        case 'Device': { const t = trucks.find(x => x[0] === s.id); return t ? [{ id: t[0], name: t[1] }] : []; }
        default: return [];
      }
    }

    let n = 0;
    const api = {
      __preview: true,
      call(method, params, ok) {
        setTimeout(() => ok(method === 'Get' ? get(params) : method === 'Add' ? 'preview' + (++n) : undefined), 250);
      },
      multiCall(calls, ok) { setTimeout(() => ok(calls.map(c => get(c[1]))), 250); },
      getSession(ok) { ok({ userName: 'preview@example.com', database: 'preview' }); }
    };

    document.getElementById('psi-preview').classList.remove('hidden');
    geotab.addin.plowSweeperInvestigation().initialize(api, {}, () => {});
  }

  // give a real host a moment to initialize first
  setTimeout(start, 500);
})();
