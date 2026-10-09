// Minimal: auto-submit the workflow picker, and poll a running job's log.
document.querySelectorAll("[data-autosubmit]").forEach(function (el) {
  el.addEventListener("change", function () { el.form.submit(); });
});

(function () {
  var pre = document.getElementById("log");
  if (!pre || pre.dataset.running !== "yes") return;
  var offset = parseInt(pre.dataset.offset, 10) || 0;
  function poll() {
    fetch("/jobs/" + pre.dataset.job + "/log?offset=" + offset, { credentials: "same-origin" })
      .then(function (r) { return r.json(); })
      .then(function (d) {
        if (d.text) { pre.textContent += d.text; }   // textContent: log text is never parsed as HTML
        offset = d.offset;
        if (d.status === "running") { setTimeout(poll, 1000); } else { location.reload(); }   // the reload shows the final status
      })
      .catch(function () { setTimeout(poll, 3000); });
  }
  poll();
})();
