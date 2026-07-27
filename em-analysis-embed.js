/* ============================================================================
   Equity & Markets Insight — analysis-page enhancer
   ----------------------------------------------------------------------------
   Progressive enhancement for every *-blog-page.html. Adds only chrome and
   reading aids — it never rewrites, reorders or removes article content,
   figures, tables or charts.

     1. ?embed=1  -> chrome-less mode + continuous height postMessage, so the
                     Base44 app can render the article at its natural height
                     instead of a fixed-height iframe with its own scrollbar.
     2. site nav + footer matching front-page.html (standalone viewing only)
     3. reading-progress bar
     4. "In this analysis" contents list built from the article's own <h2>s
     5. horizontal scroll shell around any table that lacks one (mobile)
     6. back-to-top control

   Created 2026-07-27. Companion stylesheet: em-analysis-theme.css
   ========================================================================== */
(function () {
  "use strict";

  var doc = document;
  var params = new URLSearchParams(location.search);
  var EMBEDDED = params.has("embed") || window.parent !== window;
  var SLUG = location.pathname.split("/").pop().replace(/\.html$/, "");
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

  function meta(name) {
    var m = doc.querySelector('meta[name="' + name + '"]');
    return m ? m.getAttribute("content") : null;
  }

  /* --- 1. Embed mode: tell the parent how tall we really are -------------- */
  function heightReporter() {
    var last = 0;
    function measure() {
      return Math.ceil(
        Math.max(
          doc.documentElement.scrollHeight,
          doc.body ? doc.body.scrollHeight : 0,
          doc.documentElement.offsetHeight
        )
      );
    }
    function post(force) {
      var h = measure();
      if (!force && Math.abs(h - last) < 2) return;
      last = h;
      try {
        window.parent.postMessage(
          { type: "em-article-height", slug: SLUG, height: h },
          "*"
        );
      } catch (e) { /* nothing we can do */ }
    }
    var queued = false;
    function schedule() {
      if (queued) return;
      queued = true;
      requestAnimationFrame(function () { queued = false; post(false); });
    }

    post(true);
    window.addEventListener("load", function () { post(true); });
    window.addEventListener("resize", schedule);
    if (doc.fonts && doc.fonts.ready) doc.fonts.ready.then(function () { post(true); });

    // Chart.js canvases and lazy images settle after load
    [120, 400, 900, 1800, 3000].forEach(function (ms) {
      setTimeout(function () { post(true); }, ms);
    });

    if (window.ResizeObserver) {
      new ResizeObserver(schedule).observe(doc.documentElement);
    }
    new MutationObserver(schedule).observe(doc.body, {
      childList: true, subtree: true, attributes: true, characterData: true
    });

    // let the parent ask
    window.addEventListener("message", function (ev) {
      if (ev.data && ev.data.type === "em-request-height") post(true);
    });
  }

  /* --- 2. Site chrome ----------------------------------------------------- */
  function buildNav() {
    var nav = el("div", "em-nav");
    nav.innerHTML =
      '<div class="em-nav-inner">' +
        '<a class="em-brand" href="front-page.html">' +
          '<img src="emi-logo.png" alt="Equity &amp; Markets Insight logo">' +
          '<span class="em-brand-name">Equity <b>&amp;</b> Markets Insight</span>' +
        "</a>" +
        '<nav class="em-nav-links">' +
          '<a target="_top" href="' + SITE + '/Companies">Company Analysis</a>' +
          '<a target="_top" href="' + SITE + '/Articles">Macro</a>' +
          '<a target="_top" href="' + SITE + '/company-data">Company Data</a>' +
          '<a target="_top" href="' + SITE + '/TradingLog">Trading Log</a>' +
          '<a target="_top" href="' + SITE + '/LongReads">Long Reads</a>' +
        "</nav>" +
      "</div>";
    doc.body.insertBefore(nav, doc.body.firstChild);
  }

  function buildFooter() {
    var foot = el("footer", "em-foot");
    foot.innerHTML =
      '<div class="em-foot-inner">' +
        '<div class="em-foot-grid">' +
          "<div>" +
            '<div class="em-foot-brand"><img src="emi-logo.png" alt="">' +
            "<span>Equity <b>&amp;</b> Markets Insight</span></div>" +
            '<p class="em-blurb">Independent UK equity &amp; macro research. ' +
            "Facts dated and sourced; opinion kept honestly apart.</p>" +
          "</div>" +
          '<div class="em-foot-cols">' +
            '<div class="em-foot-col"><h6>Sections</h6>' +
              '<a target="_top" href="' + SITE + '/Companies">Company Analysis</a>' +
              '<a target="_top" href="' + SITE + '/Articles">Macro</a>' +
              '<a target="_top" href="' + SITE + '/InterestRates">Rates</a>' +
              '<a target="_top" href="' + SITE + '/TradingLog">Trading Log</a>' +
              '<a target="_top" href="' + SITE + '/LongReads">Long Reads</a></div>' +
            '<div class="em-foot-col"><h6>Data</h6>' +
              '<a target="_top" href="' + SITE + '/company-data">Company Data</a>' +
              '<a href="financial-calendar.html">Financial Calendar</a>' +
              '<a href="guidance-tracker.html">Guidance Tracker</a>' +
              '<a href="rns-news.html">RNS News</a></div>' +
            '<div class="em-foot-col"><h6>Site</h6>' +
              '<a href="front-page.html">Front Page</a>' +
              '<a target="_top" href="' + SITE + '/About">About</a></div>' +
          "</div>" +
        "</div>" +
        '<p class="em-foot-note">&copy; Markets &amp; Equities Research. This page is for ' +
        "information and education only and is not financial advice. I am not a financial " +
        "adviser. Investing involves risk, including loss of capital. Do your own research " +
        "and consider seeking independent advice.</p>" +
      "</div>";
    doc.body.appendChild(foot);
  }

  /* --- 3. Hero kicker ----------------------------------------------------- */
  function buildHeroKicker() {
    var inner = doc.querySelector(".hero-banner .hero-inner");
    if (!inner || inner.querySelector(".em-hero-kicker")) return;

    var cat = meta("tile-cat") || "Company Analysis";
    var dateEl = doc.querySelector(".meta-date");
    // reuse the article's own dateline text, first clause only — never re-format
    var dateText = "";
    if (dateEl) dateText = (dateEl.textContent || "").split("·")[0].trim();

    var k = el("div", "em-hero-kicker");
    k.innerHTML =
      "<span>" + cat + "</span>" +
      '<span class="em-rule"></span>' +
      '<span class="em-date">' + dateText + "</span>";
    inner.insertBefore(k, inner.firstChild);
  }

  /* --- 4. Contents list --------------------------------------------------- */
  function slugify(s) {
    return "s-" + s.toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "").slice(0, 60);
  }

  function buildToc() {
    var container = doc.querySelector(".article-container");
    var bar = doc.querySelector(".meta-bar");
    if (!container || !bar) return;

    var heads = Array.prototype.slice.call(container.querySelectorAll("h2"));
    if (heads.length < 4) return; // not worth it on short pieces

    var list = el("div", "em-toc-list");
    var seen = {};
    heads.forEach(function (h) {
      var text = (h.textContent || "").trim();
      if (!text) return;
      var id = h.id;
      if (!id) {
        id = slugify(text);
        while (seen[id]) id += "x";
        h.id = id;
      }
      seen[id] = true;
      var a = doc.createElement("a");
      a.href = "#" + id;
      a.textContent = text;
      list.appendChild(a);
    });
    if (!list.children.length) return;

    var toc = el("details", "em-toc");
    var sum = doc.createElement("summary");
    sum.textContent = "In this analysis";
    toc.appendChild(sum);
    toc.appendChild(list);
    bar.parentNode.insertBefore(toc, bar.nextSibling);
  }

  /* --- 5. Table scroll shells --------------------------------------------- */
  function wrapTables() {
    Array.prototype.slice.call(doc.querySelectorAll(".article-container table")).forEach(function (t) {
      var p = t.parentNode;
      if (!p) return;
      var style = p.getAttribute ? (p.getAttribute("style") || "") : "";
      if (p.classList && p.classList.contains("em-table-scroll")) return;
      if (/overflow-x\s*:\s*auto/i.test(style)) return; // already wrapped
      var shell = el("div", "em-table-scroll");
      p.insertBefore(shell, t);
      shell.appendChild(t);
    });
  }

  /* --- 6. Progress bar + back to top -------------------------------------- */
  function readingAids() {
    var bar = el("div", "em-progress");
    doc.body.appendChild(bar);

    var top = doc.createElement("button");
    top.className = "em-top";
    top.type = "button";
    top.setAttribute("aria-label", "Back to top");
    top.innerHTML = "&uarr;";
    top.addEventListener("click", function () {
      window.scrollTo({ top: 0, behavior: "smooth" });
    });
    doc.body.appendChild(top);

    var ticking = false;
    function onScroll() {
      if (ticking) return;
      ticking = true;
      requestAnimationFrame(function () {
        ticking = false;
        var h = doc.documentElement.scrollHeight - window.innerHeight;
        var pct = h > 0 ? Math.min(100, (window.scrollY / h) * 100) : 0;
        bar.style.width = pct + "%";
        top.classList.toggle("show", window.scrollY > 700);
      });
    }
    window.addEventListener("scroll", onScroll, { passive: true });
    onScroll();
  }

  /* --- run ---------------------------------------------------------------- */
  ready(function () {
    if (EMBEDDED) {
      doc.body.classList.add("em-embedded");
    } else {
      buildNav();
      buildFooter();
      readingAids();
    }
    buildHeroKicker();
    buildToc();
    wrapTables();
    if (EMBEDDED) heightReporter();
  });
})();
