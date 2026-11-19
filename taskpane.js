/* Volet du complément : lit le mail, le corrige, réécrit le corps. */
(function () {
  "use strict";

  var originalHtml = null; // corps avant la dernière correction, pour « Annuler »
  var $ = function (id) { return document.getElementById(id); };

  // ---- Réglages (mémorisés dans le navigateur d'Outlook) ----
  function load(key, def) { try { var v = localStorage.getItem("correcteur." + key); return v === null ? def : JSON.parse(v); } catch (e) { return def; } }
  function save(key, val) { try { localStorage.setItem("correcteur." + key, JSON.stringify(val)); } catch (e) {} }

  var settings = {
    language: load("language", "fr"),
    autorun: load("autorun", true),
    dictionary: load("dictionary", []),
  };

  function status(msg, kind) {
    var el = $("status");
    el.textContent = msg;
    el.className = "status" + (kind ? " " + kind : "");
  }

  // ---- Accès au corps du mail (Office.js, en promesses) ----
  function getBody() {
    return new Promise(function (resolve, reject) {
      Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, function (r) {
        r.status === Office.AsyncResultStatus.Succeeded ? resolve(r.value) : reject(r.error);
      });
    });
  }
  function setBody(html) {
    return new Promise(function (resolve, reject) {
      Office.context.mailbox.item.body.setAsync(html, { coercionType: Office.CoercionType.Html }, function (r) {
        r.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(r.error);
      });
    });
  }

  // ---- Correction ----
  async function run() {
    var btn = $("btn-correct");
    btn.disabled = true;
    $("fixes").innerHTML = "";
    status("Correction en cours…");
    try {
      var html = await getBody();
      var result = await Correcteur.correctHtml(html, { language: settings.language, dictionary: settings.dictionary });
      if (!result.fixes.length) {
        status("Aucune faute trouvée.", "ok");
        return;
      }
      originalHtml = html;
      await setBody(result.html);
      var n = result.fixes.length;
      status(n + (n > 1 ? " fautes corrigées." : " faute corrigée."), "ok");
      showFixes(result.fixes);
      $("btn-undo").hidden = false;
    } catch (e) {
      status((e && e.message) || "La correction a échoué.", "err");
    } finally {
      btn.disabled = false;
    }
  }

  async function undo() {
    if (originalHtml === null) return;
    // confirm() est bloqué dans Outlook : on demande un second clic.
    var btn = $("btn-undo");
    if (!btn.dataset.armed) {
      btn.dataset.armed = "1";
      btn.textContent = "Confirmer : ce que vous avez tapé depuis sera perdu";
      setTimeout(function () { delete btn.dataset.armed; btn.textContent = "Annuler la correction"; }, 5000);
      return;
    }
    delete btn.dataset.armed;
    btn.textContent = "Annuler la correction";
    try {
      await setBody(originalHtml);
      originalHtml = null;
      $("btn-undo").hidden = true;
      $("fixes").innerHTML = "";
      status("Correction annulée.", "ok");
    } catch (e) {
      status("Impossible d'annuler : " + ((e && e.message) || "erreur"), "err");
    }
  }

  function showFixes(fixes) {
    var ul = $("fixes");
    fixes.forEach(function (f) {
      var li = document.createElement("li");
      var pair = document.createElement("div");
      pair.className = "pair";
      var o = document.createElement("span"); o.className = "old"; o.textContent = f.original;
      var n = document.createElement("span"); n.className = "new"; n.textContent = f.replacement;
      pair.append(o, " → ", n);
      var msg = document.createElement("div"); msg.className = "msg"; msg.textContent = f.message;
      li.append(pair, msg);
      if (f.type === "misspelling") {
        var b = document.createElement("button");
        b.className = "link";
        b.textContent = "Ne plus corriger « " + f.original + " »";
        b.onclick = function () { addWord(f.original); b.replaceWith("Ajouté. Cliquez sur « Annuler la correction » puis « Corriger » pour rétablir ce mot."); };
        li.append(b);
      }
      ul.append(li);
    });
  }

  // ---- Dictionnaire personnel ----
  function addWord(w) {
    w = (w || "").trim();
    if (!w || settings.dictionary.some(function (x) { return x.toLowerCase() === w.toLowerCase(); })) return;
    settings.dictionary.push(w);
    settings.dictionary.sort(function (a, b) { return a.localeCompare(b, "fr"); });
    save("dictionary", settings.dictionary);
    renderDict();
  }
  function renderDict() {
    var ul = $("dict-list");
    ul.innerHTML = "";
    if (!settings.dictionary.length) { var e = document.createElement("li"); e.textContent = "Aucun mot."; e.style.color = "var(--muted)"; ul.append(e); return; }
    settings.dictionary.forEach(function (w, i) {
      var li = document.createElement("li");
      var s = document.createElement("span"); s.textContent = w;
      var b = document.createElement("button"); b.title = "Retirer"; b.textContent = "✕";
      b.onclick = function () { settings.dictionary.splice(i, 1); save("dictionary", settings.dictionary); renderDict(); };
      li.append(s, b);
      ul.append(li);
    });
  }

  // ---- Démarrage ----
  Office.onReady(function (info) {
    $("language").value = settings.language;
    $("autorun").checked = settings.autorun;
    $("language").onchange = function (e) { settings.language = e.target.value; save("language", settings.language); };
    $("autorun").onchange = function (e) { settings.autorun = e.target.checked; save("autorun", settings.autorun); };
    $("dict-form").onsubmit = function (e) { e.preventDefault(); addWord($("dict-input").value); $("dict-input").value = ""; };
    $("btn-correct").onclick = run;
    $("btn-undo").onclick = undo;
    renderDict();

    if (info.host !== Office.HostType.Outlook || !Office.context.mailbox || !Office.context.mailbox.item) {
      status("Ouvrez ce volet depuis un mail en cours de rédaction.", "err");
      $("btn-correct").disabled = true;
      return;
    }
    if (settings.autorun) run();
  });
})();
