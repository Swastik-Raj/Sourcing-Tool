// Minimal: auto-submit the workflow picker, the product picker on the Search page, and polling a running job's log.
document.querySelectorAll("[data-autosubmit]").forEach(function (el) {
  el.addEventListener("change", function () { el.form.submit(); });
});

// Search page: filter the product checklist and show how many products are selected (the server recounts on submit).
// Nothing is pre-selected, so "Next: check the cost" stays greyed out, with its reason beside it, until a choice is made.
// Without JavaScript the button works and the server refuses an empty choice with the same reason.
(function () {
  var form = document.getElementById("search-form");
  if (!form) return;
  var out = document.getElementById("sel-count"), filter = document.getElementById("sku-filter");
  var next = form.querySelector("button[data-needs-choice]"), first = document.getElementById("choose-first");
  function count() {
    if (!out) return;                      // the sheet preview form has no counter: the next page counts
    var choice = form.querySelector("input[name=choice]:checked");
    var n = !choice ? 0 : choice.value === "choose" ? form.querySelectorAll("input[name=sku]:checked").length : choice.dataset.count;
    out.textContent = choice ? n + " product" + (String(n) === "1" ? "" : "s") + " selected." : "";
    if (next && !next.hasAttribute("data-blocked")) {     // a reason the server gave (setup missing, another task) is left alone
      next.disabled = String(n) === "0";
      if (first) { first.hidden = !next.disabled; }
    }
  }
  form.addEventListener("change", function (e) {
    var pick = form.querySelector("input[value=choose], input[value=pick]");
    if (e.target.name === "sku" && e.target.checked && pick) { pick.checked = true; }
    if (e.target.name === "limit" && e.target.value && form.querySelector("input[value=first]")) { form.querySelector("input[value=first]").checked = true; }
    count();
  });
  if (filter) {
    filter.addEventListener("input", function () {
      var q = filter.value.trim().toLowerCase();
      document.querySelectorAll("#sku-list li").forEach(function (li) { li.hidden = q !== "" && li.textContent.toLowerCase().indexOf(q) === -1; });
    });
  }
  count();
})();

// Email sellers: "Copy the text" beside a read-only box. The button stays hidden without JavaScript (the text can still be
// selected by hand). Where the clipboard cannot be written (not localhost or https), the text is selected for Ctrl+C.
document.querySelectorAll("button[data-copy]").forEach(function (btn) {
  var box = document.getElementById(btn.dataset.copy), said = document.getElementById(btn.dataset.copy + "-status");
  if (!box) return;
  btn.hidden = false;
  btn.addEventListener("click", function () {
    function say(text) { if (said) { said.textContent = text; } }   // textContent only
    box.focus();
    box.select();
    if (navigator.clipboard && window.isSecureContext) {
      navigator.clipboard.writeText(box.value).then(function () { say("Copied."); }, function () { say("Selected: press Ctrl+C to copy."); });
    } else {
      say("Selected: press Ctrl+C to copy.");
    }
  });
});

(function () {
  var pre = document.getElementById("log");
  if (!pre || pre.dataset.running !== "yes") return;
  var offset = parseInt(pre.dataset.offset, 10) || 0;
  function set(id, value) { var el = document.getElementById(id); if (el) { el.textContent = value; } }   // textContent only
  function showProgress(p) {
    set("p-started", p.started); set("p-total", p.total || "?"); set("p-done", p.done);
    set("p-current", (p.current || []).join(", ") || "starting");
    set("p-cost", Number(p.cost || 0).toFixed(4));
    set("p-elapsed", Math.floor(p.elapsed_s / 60) + " min " + (p.elapsed_s % 60) + " s");
    var bar = document.getElementById("p-bar");
    if (bar) { bar.max = p.total || 1; bar.value = p.done || 0; }
  }
  function poll() {
    fetch("/jobs/" + pre.dataset.job + "/log?offset=" + offset, { credentials: "same-origin" })
      .then(function (r) { return r.json(); })
      .then(function (d) {
        if (d.text) { pre.textContent += d.text; }   // textContent: log text is never parsed as HTML
        offset = d.offset;
        if (d.progress) { showProgress(d.progress); }
        if (d.status === "running") { setTimeout(poll, 1000); } else { location.reload(); }   // the reload shows the final status
      })
      .catch(function () { setTimeout(poll, 3000); });
  }
  poll();
})();
