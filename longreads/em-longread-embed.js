/* ============================================================================
   Equity & Markets Insight — long-read chrome
   ----------------------------------------------------------------------------
   Progressive enhancement for every /longreads/*.html page. Adds only site
   chrome (header + footer) — it never rewrites, reorders or removes an
   article's own headline, body copy, stats, charts, tables or sources.

     1. Header: two of the four long reads already ship a bespoke
        .site-header (styled by em-longread-theme.css); this builds the same
        markup for the two that currently have none, and upgrades the
        hand-drawn placeholder mark in the older ones to the real logo file
        so every long read carries an identical, on-brand header.
     2. Footer: appends the site's standard footer (colours/type only —
        namespaced em-foot-*, not shared with ../em-analysis-theme.css) if
        one is not already present.

   Deliberately does NOT touch the page's own progress bar / TOC / reveal /
   Chart.js logic already wired up in each article's bottom <script> — that
   stays exactly as authored.

   Created 2026-07-27. Companion stylesheet: em-longread-theme.css
   ========================================================================== */
(function () {
  "use strict";

  var doc = document;
  var SITE = "https://equityandmarketsinsight.com";

  function ready(fn) {
    if (doc.readyState === "loading") doc.addEventListener("DOMContentLoaded", fn);
    else fn();
  }

  function el(tag, cls, html) {
    var n = doc.createElement(tag);
    if (cls) n.className = cls;
    if (html != null) n.innerHTML = html;
    return n;
  }

  /* --- 1. Header ------------------------------------------------------------ */
  function navLinks(activeHref) {
    var items = [
      ["Home", "../front-page.html"],
      ["Company Analysis", SITE + "/Companies"],
      ["Macro", SITE + "/Articles"],
      ["Company Data", SITE + "/company-data"],
      ["Trading Log", SITE + "/TradingLog"],
      ["Long Reads", SITE + "/LongReads"]
    ];
    return items.map(function (item) {
      var active = item[0] === "Long Reads" ? ' class="active"' : "";
      return '<a href="' + item[1] + '"' + active + '>' + item[0] + "</a>";
    }).join("");
  }

  function ensureHeader() {
    var header = doc.querySelector(".site-header");

    if (!header) {
      header = el("header", "site-header");
      header.innerHTML =
        '<a class="brand" href="../front-page.html">' +
          '<span class="brand-icon"><img src="../emi-logo.png" alt="Equity &amp; Markets Insight logo"></span>' +
          '<span class="brand-name">Equity &amp; Markets Insight</span>' +
        "</a>" +
        "<nav>" + navLinks() + "</nav>";
      doc.body.insertBefore(header, doc.body.firstChild);
      return;
    }

    // Already present (the older long reads) — upgrade the hand-drawn SVG
    // mark to the real logo file; leave everything else in the existing
    // header untouched.
    var icon = header.querySelector(".brand-icon");
    if (icon && !icon.querySelector("img")) {
      icon.innerHTML = '<img src="../emi-logo.png" alt="Equity &amp; Markets Insight logo">';
    }

    // Mark "Long Reads" active without assuming the existing nav already
    // has that link — replace the nav content wholesale with the current
    // site nav so every long read carries the same set of links.
    var nav = header.querySelector("nav");
    if (nav) nav.innerHTML = navLinks();
  }

  /* --- 2. Footer -------------------------------------------------------------- */
  function buildFooter() {
    if (doc.querySelector(".em-foot")) return;
    var foot = el("footer", "em-foot");
    foot.innerHTML =
      '<div class="em-foot-inner">' +
        '<div class="em-foot-grid">' +
          "<div>" +
            '<div class="em-foot-brand"><img src="../emi-logo.png" alt="">' +
            "<span>Equity <b>&amp;</b> Markets Insight</span></div>" +
            '<p class="em-blurb">Independent UK equity &amp; macro research. ' +
            "Facts dated and sourced; opinion kept honestly apart.</p>" +
          "</div>" +
          '<div class="em-foot-cols">' +
            '<div class="em-foot-col"><h6>Sections</h6>' +
              '<a href="' + SITE + '/Companies">Company Analysis</a>' +
              '<a href="' + SITE + '/Articles">Macro</a>' +
              '<a href="' + SITE + '/InterestRates">Rates</a>' +
              '<a href="' + SITE + '/TradingLog">Trading Log</a>' +
              '<a href="' + SITE + '/LongReads">Long Reads</a></div>' +
            '<div class="em-foot-col"><h6>Data</h6>' +
              '<a href="' + SITE + '/company-data">Company Data</a>' +
              '<a href="../financial-calendar.html">Financial Calendar</a>' +
              '<a href="../rns-news.html">RNS News</a></div>' +
            '<div class="em-foot-col"><h6>Site</h6>' +
              '<a href="../front-page.html">Front Page</a>' +
              '<a href="' + SITE + '/About">About</a></div>' +
          "</div>" +
        "</div>" +
        '<p class="em-foot-note">&copy; Markets &amp; Equities Research. This page is for ' +
        "information and education only and is not financial advice. I am not a financial " +
        "adviser. Investing involves risk, including loss of capital. Do your own research " +
        "and consider seeking independent advice.</p>" +
      "</div>";
    doc.body.appendChild(foot);
  }

  ready(function () {
    ensureHeader();
    buildFooter();
  });
})();
