// Company Data access gate (24 November 2026).
//
// Free readers get a limited number of company look-ups a day; Premium readers
// get unlimited use. The limit, the count and the reader's plan all live in the
// site (equityandmarketsinsight.com, the Base44 app), not here. Before app.js
// opens a company it asks the site, which frames this page:
//
//   this page -> site   { type: 'emi-cd-lookup', id, key }
//   site -> this page   { type: 'emi-cd-verdict', id, allowed, ... }
//
// A verdict counts only when it comes from the parent window AND from one of the
// site's origins. Opened on its own (the raw github.io address, or framed by any
// other site) there is nobody to answer, so nothing beyond the sample company
// opens and the reader is pointed at the site. That is what stops the limit being
// skipped by opening this address directly.
//
// This is a display gate. data.js is a public file; a determined reader can still
// read it. No limit number is kept here on purpose: the site owns it (plans.js).
(function () {
  'use strict';

  var SITE = 'https://equityandmarketsinsight.com';
  var SHELL_ORIGINS = [SITE, 'https://www.equityandmarketsinsight.com',
    'https://perky-equity-insight-hub.base44.app'];
  var TIMEOUT_MS = 10000;

  var framed = window.parent !== window;
  var pending = {};
  var seq = 0;

  function shellTargets() {
    if (!framed) return [];
    // The parent's origin where the browser says (Chrome, Safari, Edge); otherwise
    // every known site origin - the browser drops the ones that do not match.
    try {
      var ao = window.location.ancestorOrigins;
      if (ao && ao.length) return SHELL_ORIGINS.indexOf(ao[0]) === -1 ? [] : [ao[0]];
    } catch (e) {}
    return SHELL_ORIGINS;
  }

  window.addEventListener('message', function (event) {
    if (SHELL_ORIGINS.indexOf(event.origin) === -1) return;
    if (event.source !== window.parent) return;
    var d = event.data;
    if (!d || d.type !== 'emi-cd-verdict') return;
    var p = pending[d.id];
    if (!p) return;
    delete pending[d.id];
    clearTimeout(p.timer);
    p.resolve(d.allowed === true ? { allowed: true } : { allowed: false, reason: 'limit' });
  });

  // Resolves to { allowed: true } or { allowed: false, reason }, where reason is
  // 'limit' (the site said no), 'outside' (not inside the site) or 'no-answer'.
  function request(key) {
    return new Promise(function (resolve) {
      var targets = shellTargets();
      if (!targets.length) { resolve({ allowed: false, reason: 'outside' }); return; }
      var id = 'cd-' + Date.now() + '-' + (++seq);
      var timer = setTimeout(function () {
        delete pending[id];
        resolve({ allowed: false, reason: 'no-answer' });
      }, TIMEOUT_MS);
      pending[id] = { resolve: resolve, timer: timer };
      targets.forEach(function (origin) {
        try { window.parent.postMessage({ type: 'emi-cd-lookup', id: id, key: key }, origin); } catch (e) {}
      });
    });
  }

  window.EMI_COMPANY_GATE = { request: request, site: SITE, pageUrl: SITE + '/company-data' };
})();
