/* Plow/Sweeper Activity Investigation - MyGeotab add-in (v1) */
geotab.addin.plowSweeperInvestigation = function () {
  'use strict';

  const ADDIN_ID = 'aGFHyz5HvRWaByIqrNlLgdQ';   // fixed AddInData key for this add-in
  const PREFIX = 'PSI-';                          // name prefix for zones and rules we create
  const MAX_MONTHS = 6;

  let api, map, drawn, lastResults = [];
  const $ = id => document.getElementById(id);

  // ---------- API helpers ----------
  const call = (method, params) => new Promise((ok, fail) => api.call(method, params, ok, fail));
  const multi = calls => new Promise((ok, fail) => api.multiCall(calls, ok, fail));
  const errText = e => (e && (e.message || e.name)) || String(e);
  const esc = s => String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

  // ---------- Rule condition builder (encoding confirmed against Newmarket Plow Down rule) ----------
  const seq = i => String(i).padStart(16, '0');
  const aux = n => ({ conditionType: 'Aux' + n, value: 1 });
  const diagEq = (diag, v) => ({
    conditionType: 'IsValueEqualTo', value: v,
    children: [{ conditionType: 'FilterStatusDataByDiagnostic', sequence: seq(0), diagnostic: { id: diag } }]
  });
  const PRESETS = {
    plow: () => [aux(6), aux(7), diagEq('DiagnosticThirdPartyAux6Id', 1), diagEq('DiagnosticThirdPartyAux7Id', 1)],
    spreader: () => [diagEq('DiagnosticControllerStatusId', 4), diagEq('DiagnosticControllerStatusId', 6),
                     diagEq('DiagnosticControllerStatusId', 7), aux(8)],
    sander: () => [aux(2)],
    sweeper: () => [aux(3)]
  };
  const PRESET_LABEL = { plow: 'Plow', spreader: 'Spreader', sander: 'Sander', sweeper: 'Sweeper' };

  function equipmentConditions(presets, auxes) {
    const list = [], seen = new Set();
    const add = c => { const k = JSON.stringify(c); if (!seen.has(k)) { seen.add(k); list.push(c); } };
    presets.forEach(p => PRESETS[p]().forEach(add));
    auxes.forEach(n => add(aux(n)));
    return list.map((c, i) => Object.assign({}, c, { sequence: seq(i) }));
  }

  function buildCondition(zoneId, equip) {
    const top = [{ conditionType: 'InsideArea', zone: { id: zoneId }, sequence: seq(0) }];
    if (equip.length) {
      top.push({ conditionType: 'DurationLongerThan', value: 1, unit: 's', sequence: seq(1),
        children: [{ conditionType: 'Or', sequence: seq(0), children: equip }] });
      top.push({ conditionType: 'Ignition', value: 1, sequence: seq(2) });
    }
    return { conditionType: 'And', sequence: seq(0), children: top };
  }

  // ---------- Dates and formatting ----------
  const localMidnightISO = d => new Date(d + 'T00:00:00').toISOString();
  const localEndOfDayISO = d => new Date(d + 'T23:59:59').toISOString();
  const fmt = iso => new Date(iso).toLocaleString([], { dateStyle: 'medium', timeStyle: 'short' });
  const fmtDate = iso => new Date(iso).toLocaleDateString([], { dateStyle: 'medium' });
  const dur = ms => { const m = Math.round(ms / 60000); return m < 60 ? `${m} min` : `${Math.floor(m / 60)} h ${m % 60} min`; };

  function validate(name, from, to) {
    if (!name) return 'Enter a run name.';
    if (!from || !to) return 'Choose a start and end date.';
    const f = new Date(from + 'T00:00:00'), t = new Date(to + 'T00:00:00'), today = new Date();
    const min = new Date(); min.setMonth(min.getMonth() - MAX_MONTHS);
    if (f > today) return 'Start date cannot be in the future.';
    if (f < min) return `Start date cannot be more than ${MAX_MONTHS} months ago.`;
    if (t < f) return 'End date must be on or after the start date.';
    if (!drawn) return 'Draw the area on the map.';
    return null;
  }

  function msg(el, text, kind) {
    $(el).innerHTML = text ? `<div class="msg ${kind || ''}">${text}</div>` : '';
  }

  // ---------- Map ----------
  // MyGeotab drops the page <head>, so load the map stylesheets from here.
  const CSS = [
    ['https://cdn.jsdelivr.net/npm/leaflet@1.9.4/dist/leaflet.css',
     'sha384-sHL9NAb7lN7rfvG5lfHpm643Xkcjzp4jFvuavGOndn6pjVqS6ny56CAt3nsEVT4H'],
    ['https://cdn.jsdelivr.net/npm/leaflet-draw@1.0.4/dist/leaflet.draw.css',
     'sha384-NZLkVuBRMEeB4VeZz27WwTRvlhec30biQ8Xx7zG7JJnkvEKRg5qi6BNbEXo9ydwv']
  ];
  function ensureCss() {
    CSS.forEach(([href, sri]) => {
      if (document.querySelector(`link[href="${href}"]`)) return;
      const l = document.createElement('link');
      l.rel = 'stylesheet'; l.href = href; l.integrity = sri; l.crossOrigin = 'anonymous';
      document.head.appendChild(l);
    });
  }

  // Starting view: the user's last view, else the fleet's area, else a wide North America view.
  const VIEW_KEY = 'psi-map-view';
  const loadView = () => { try { return JSON.parse(localStorage.getItem(VIEW_KEY)); } catch (e) { return null; } };
  const saveView = () => {
    try { const c = map.getCenter(); localStorage.setItem(VIEW_KEY, JSON.stringify({ lat: c.lat, lng: c.lng, zoom: map.getZoom() })); } catch (e) {}
  };

  async function fitToFleet() {
    try {
      const pts = (await call('Get', { typeName: 'DeviceStatusInfo',
        propertySelector: { fields: ['latitude', 'longitude'] } }))
        .filter(s => s.latitude && s.longitude);
      if (!pts.length) return;
      // ignore outliers: keep vehicles within ~150 km of the fleet's median position
      const median = arr => { const s = arr.slice().sort((a, b) => a - b); return s[Math.floor(s.length / 2)]; };
      const mLat = median(pts.map(p => p.latitude)), mLng = median(pts.map(p => p.longitude));
      const near = pts.filter(p => Math.abs(p.latitude - mLat) < 1.35 && Math.abs(p.longitude - mLng) < 1.9);
      const bounds = L.latLngBounds((near.length ? near : pts).map(p => [p.latitude, p.longitude]));
      map.fitBounds(bounds.pad(0.2), { maxZoom: 13 });
    } catch (e) { /* keep the fallback view */ }
  }

  function initMap() {
    if (map) return;
    ensureCss();
    const saved = loadView();
    map = L.map('psi-map').setView(saved ? [saved.lat, saved.lng] : [50, -95], saved ? saved.zoom : 4);
    map.on('moveend', saveView);
    if (!saved) fitToFleet();
    L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png',
      { maxZoom: 19, attribution: '&copy; OpenStreetMap contributors' }).addTo(map);
    const items = new L.FeatureGroup().addTo(map);
    const style = { color: '#0078d3', weight: 2, fillOpacity: 0.15 };
    map.addControl(new L.Control.Draw({
      draw: { polygon: { allowIntersection: false, showArea: true, shapeOptions: style }, rectangle: { shapeOptions: style },
              polyline: false, circle: false, circlemarker: false, marker: false },
      edit: { featureGroup: items }
    }));
    map.on(L.Draw.Event.CREATED, e => { items.clearLayers(); items.addLayer(e.layer); drawn = e.layer; });
    map.on(L.Draw.Event.DELETED, () => { if (!items.getLayers().length) drawn = null; });
    map.on(L.Draw.Event.EDITED, () => { drawn = items.getLayers()[0] || null; });
  }

  // ---------- Address search (MyGeotab geocoder) ----------
  let searchMarker;
  const hint = (text, isError) => {
    $('psi-map-hint').textContent = text;
    $('psi-map-hint').classList.toggle('error', !!isError);
  };
  async function searchAddress(ev) {
    ev.preventDefault();
    const q = $('psi-address').value.trim();
    if (!q) return;
    hint('Searching...');
    try {
      const res = await call('GetCoordinates', { addresses: [q] });
      const c = res && res[0];
      if (!c || (c.x === 0 && c.y === 0)) return hint(`No match found for "${q}". Try adding the city and province or state.`, true);
      map.setView([c.y, c.x], 16);
      if (searchMarker) searchMarker.remove();
      searchMarker = L.circleMarker([c.y, c.x], { radius: 8, color: '#fff', weight: 2, fillColor: '#0078d3', fillOpacity: 1 })
        .addTo(map).bindTooltip(q).openTooltip();
      hint('Found it. Now draw the area with the polygon or rectangle tool (top left of the map).');
    } catch (e) {
      hint('Address search failed: ' + errText(e), true);
    }
  }

  function zonePoints() {
    const pts = drawn.getLatLngs()[0].map(p => ({ x: p.lng, y: p.lat }));
    pts.push(pts[0]);   // close the polygon
    return pts;
  }

  // ---------- Submit ----------
  async function submit() {
    const name = $('psi-name').value.trim(), from = $('psi-from').value, to = $('psi-to').value;
    const err = validate(name, from, to);
    if (err) return msg('psi-new-msg', err, 'error');

    const presets = [...document.querySelectorAll('[data-preset]:checked')].map(c => c.dataset.preset);
    const auxes = [...document.querySelectorAll('[data-aux]:checked')].map(c => Number(c.dataset.aux));
    const fromISO = localMidnightISO(from), toISO = localEndOfDayISO(to);
    const label = PREFIX + name;
    $('psi-submit').disabled = true;
    msg('psi-new-msg', 'Creating zone...');

    let zoneId, ruleId;
    try {
      zoneId = await call('Add', { typeName: 'Zone', entity: {
        name: label, points: zonePoints(), groups: [{ id: 'GroupCompanyId' }], zoneTypes: [],
        activeFrom: fromISO, activeTo: '2050-01-01T00:00:00.000Z', displayed: true, mustIdentifyStops: true,
        comment: 'Created by Plow/Sweeper Activity Investigation add-in' } });

      msg('psi-new-msg', 'Creating rule...');
      ruleId = await call('Add', { typeName: 'Rule', entity: {
        name: label, baseType: 'Custom', groups: [{ id: 'GroupCompanyId' }],
        color: { r: 0, g: 120, b: 211, a: 255 },
        comment: 'Created by Plow/Sweeper Activity Investigation add-in',
        activeFrom: fromISO, activeTo: '2050-01-01T00:00:00.000Z',
        condition: buildCondition(zoneId, equipmentConditions(presets, auxes)) } });

      msg('psi-new-msg', 'Requesting reprocess...');
      const requestId = await call('Add', { typeName: 'ReprocessRequest', entity: {
        fromDate: fromISO, ruleIds: [ruleId],
        deviceFilterType: 'Devices', deviceFilterIds: ['AllDevicesReprocessRequest'] } });

      const session = await new Promise(ok => api.getSession(ok));
      await call('Add', { typeName: 'AddInData', entity: {
        addInId: ADDIN_ID, groups: [{ id: 'GroupCompanyId' }],
        details: { runName: name, zoneId, ruleId, reprocessRequestId: requestId, fromDate: fromISO, toDate: toISO,
                   presets, aux: auxes, createdBy: session.userName, createdAt: new Date().toISOString(), cleanedUp: false } } });

      const days = (new Date(to) - new Date(from)) / 864e5 + 1;
      msg('psi-new-msg', `Submitted. ${days > 7 ? 'Ranges over 7 days process overnight. Check Recent runs tomorrow morning.'
                                        : 'Short ranges usually finish within the hour. Check Recent runs.'}`, 'ok');
    } catch (e) {
      // roll back what was created so a failed submit leaves nothing behind
      if (ruleId) await call('Remove', { typeName: 'Rule', entity: { id: ruleId } }).catch(() => {});
      if (zoneId) await call('Remove', { typeName: 'Zone', entity: { id: zoneId } }).catch(() => {});
      msg('psi-new-msg', 'Submit failed: ' + esc(errText(e)) +
        (/Terminated/.test(errText(e)) ? ' (all devices are terminated)' : ''), 'error');
    } finally {
      $('psi-submit').disabled = false;
    }
  }

  // ---------- Recent runs ----------
  async function loadRuns() {
    msg('psi-runs-msg', 'Loading...');
    $('psi-results').classList.add('hidden');
    try {
      const rows = (await call('Get', { typeName: 'AddInData', search: { addInId: ADDIN_ID } }))
        .map(r => Object.assign({ _id: r.id }, typeof r.details === 'string' ? JSON.parse(r.details) : r.details))
        .filter(r => !r.cleanedUp)
        .sort((a, b) => b.createdAt.localeCompare(a.createdAt));
      msg('psi-runs-msg', '');
      if (!rows.length) {
        $('psi-runs-list').innerHTML = '<div class="card empty">No runs yet. Submit an investigation to see it here.</div>';
        return;
      }

      const calls = [];
      rows.forEach(r => {
        calls.push(['Get', { typeName: 'ReprocessRequest', search: { id: r.reprocessRequestId } }]);
        calls.push(['Get', { typeName: 'DeviceReprocessRequest', search: { reprocessRequestId: r.reprocessRequestId },
                             propertySelector: { fields: ['status'] } }]);
      });
      const res = await multi(calls);

      $('psi-runs-list').innerHTML = rows.map((r, i) => {
        const parent = (res[2 * i] || [])[0], kids = res[2 * i + 1] || [];
        const status = parent ? parent.status : 'Unknown';
        const done = kids.filter(k => k.status === 'Success').length;
        const live = done + kids.filter(k => k.status === 'Scheduled').length;
        const pct = status === 'Complete' ? 100 : live ? Math.round(done / live * 100) : 0;
        const badge = status === 'Complete' ? ['ready', 'Ready'] : status === 'Scheduled' ? ['pending', 'Processing']
                    : status === 'Cancelled' ? ['none', 'No vehicle data'] : ['bad', status];
        const tags = [...(r.presets || []).map(p => PRESET_LABEL[p] || p), ...(r.aux || []).map(n => 'Aux ' + n)];
        return `<div class="card run">
          <div>
            <div class="run-name">${esc(r.runName)}</div>
            <div class="run-meta">${fmtDate(r.fromDate)} to ${fmtDate(r.toDate)} · submitted ${fmtDate(r.createdAt)} by ${esc(r.createdBy || '')}</div>
            <div>${(tags.length ? tags : ['Any vehicle in area']).map(t => `<span class="tag">${esc(t)}</span>`).join('')}</div>
          </div>
          <div>
            <span class="badge ${badge[0]}">${badge[1]}</span>
            ${status === 'Cancelled' ? '' : `<div class="bar"><div style="width:${pct}%"></div></div>
            <div class="bar-label">${live ? `${done} of ${live} vehicles processed` : 'Waiting to start'}</div>`}
          </div>
          <div class="actions">
            ${status === 'Complete' ? `<button class="btn" data-view="${i}">View results</button>` : ''}
            <button class="btn quiet" data-clean="${i}">Clean up</button>
          </div>
        </div>`;
      }).join('');
      document.querySelectorAll('[data-view]').forEach(b => b.onclick = () => viewResults(rows[b.dataset.view]));
      document.querySelectorAll('[data-clean]').forEach(b => b.onclick = () => cleanUp(rows[b.dataset.clean]));
    } catch (e) {
      msg('psi-runs-msg', 'Could not load runs: ' + esc(errText(e)), 'error');
    }
  }

  async function viewResults(run) {
    msg('psi-runs-msg', 'Loading results...');
    try {
      const events = await call('Get', { typeName: 'ExceptionEvent', search: {
        ruleSearch: { id: run.ruleId }, fromDate: run.fromDate, toDate: run.toDate } });
      const ids = [...new Set(events.map(e => e.device.id))];
      const devs = ids.length ? await multi(ids.map(id => ['Get', { typeName: 'Device', search: { id },
        propertySelector: { fields: ['id', 'name'] } }])) : [];
      const names = {};
      devs.forEach(d => { if (d[0]) names[d[0].id] = d[0].name; });

      lastResults = events.map(e => ({
        vehicle: names[e.device.id] || e.device.id, deviceId: e.device.id,
        start: e.activeFrom, end: e.activeTo, ms: new Date(e.activeTo) - new Date(e.activeFrom)
      })).sort((a, b) => a.start.localeCompare(b.start));

      const total = lastResults.reduce((s, r) => s + r.ms, 0);
      $('psi-results-title').textContent = run.runName;
      $('psi-stat-visits').textContent = lastResults.length;
      $('psi-stat-vehicles').textContent = ids.length;
      $('psi-stat-time').textContent = dur(total);
      $('psi-results-body').innerHTML = lastResults.length ? lastResults.map(r =>
        `<tr><td>${esc(r.vehicle)}</td><td>${fmt(r.start)}</td><td>${fmt(r.end)}</td><td>${dur(r.ms)}</td>
         <td><a href="${tripsLink(r)}" target="_top">View trips</a></td></tr>`).join('')
        : '<tr><td colspan="5" class="empty">No activity found in this area for the selected range.</td></tr>';
      $('psi-results').classList.remove('hidden');
      $('psi-results').scrollIntoView({ behavior: 'smooth' });
      msg('psi-runs-msg', '');
    } catch (e) {
      msg('psi-runs-msg', 'Could not load results: ' + esc(errText(e)), 'error');
    }
  }

  // MyGeotab trips history deep link (format to verify in the target MyGeotab version)
  function tripsLink(r) {
    return `#tripsHistory,dateRange:(endDate:'${r.end}',startDate:'${r.start}'),devices:!(${r.deviceId})`;
  }

  async function cleanUp(run) {
    if (!confirm(`Archive the rule and remove the zone for "${run.runName}"?\n\nResults will no longer be available. Export them first if you need them.`)) return;
    try {
      await call('Remove', { typeName: 'Rule', entity: { id: run.ruleId, reason: 'ManualArchived' } });
      await call('Remove', { typeName: 'Zone', entity: { id: run.zoneId } }).catch(() => {});
      const details = Object.assign({}, run, { cleanedUp: true }); delete details._id;
      await call('Set', { typeName: 'AddInData', entity: { id: run._id, addInId: ADDIN_ID,
        groups: [{ id: 'GroupCompanyId' }], details } });
      loadRuns();
    } catch (e) {
      msg('psi-runs-msg', 'Clean up failed: ' + esc(errText(e)), 'error');
    }
  }

  function exportCsv() {
    const q = s => `"${String(s).replace(/"/g, '""')}"`;
    const lines = [['Vehicle', 'Entered', 'Exited', 'Duration (min)'].join(',')]
      .concat(lastResults.map(r => [q(r.vehicle), r.start, r.end, Math.round(r.ms / 60000)].join(',')));
    const a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob([lines.join('\n')], { type: 'text/csv' }));
    a.download = ($('psi-results-title').textContent || 'results').replace(/[^\w -]/g, '') + '.csv';
    a.click();
  }

  function showTab(tab) {
    $('psi-tab-new').classList.toggle('active', tab === 'new');
    $('psi-tab-runs').classList.toggle('active', tab === 'runs');
    $('psi-new').classList.toggle('hidden', tab !== 'new');
    $('psi-runs').classList.toggle('hidden', tab !== 'runs');
    if (tab === 'new') setTimeout(() => map.invalidateSize(), 0);
    if (tab === 'runs') loadRuns();
  }

  // ---------- Add-in lifecycle ----------
  return {
    initialize(freshApi, state, callback) {
      api = freshApi;
      if (!freshApi.__preview) window.__psiLive = true;   // tells preview.js to stand down
      const today = new Date(), weekAgo = new Date(Date.now() - 6 * 864e5);
      const iso = d => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
      $('psi-to').value = iso(today);
      $('psi-from').value = iso(weekAgo);
      initMap();
      $('psi-tab-new').onclick = () => showTab('new');
      $('psi-tab-runs').onclick = () => showTab('runs');
      $('psi-submit').onclick = submit;
      $('psi-search').onsubmit = searchAddress;
      $('psi-refresh').onclick = loadRuns;
      $('psi-csv').onclick = exportCsv;
      callback();
    },
    focus(freshApi) {
      api = freshApi;
      setTimeout(() => map.invalidateSize(), 0);
    },
    blur() {}
  };
};
