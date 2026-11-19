/*
 * Correcteur — logique de correction, indépendante d'Outlook.
 *
 * Principe :
 *  1. On lit le HTML du mail et on en extrait le texte, morceau par morceau
 *     (un « segment » = un nœud texte du HTML), en s'arrêtant à la signature
 *     et à l'historique de la conversation (on ne corrige que ce qu'on écrit).
 *  2. On envoie ce texte à LanguageTool, qui renvoie les fautes trouvées.
 *  3. On garde seulement l'orthographe et la grammaire (pas le style), et on
 *     écarte ce qui ressemble à un nom propre, une adresse, un mot du dictionnaire perso.
 *  4. On remplace chaque faute dans le nœud texte d'origine : la mise en forme
 *     (gras, liens, couleurs…) n'est pas touchée.
 */
(function (global) {
  "use strict";

  var API_URL = "https://api.languagetool.org/v2/check";
  var MAX_CHARS = 15000; // l'API gratuite accepte ~20 Ko par requête

  // Début de l'historique (réponse / transfert) : tout ce qui suit est ignoré.
  var THREAD_MARKERS = /^(appendonsend|divRplyFwdMsg|mail-editor-reference-message-container|x_divRplyFwdMsg)$/i;
  // Blocs à sauter entièrement (signature, citations).
  var SKIP_ID = /signature/i;
  var SKIP_CLASS = /(gmail_quote|gmail_signature|moz-signature|elementToProof-signature)/i;
  var SKIP_TAGS = { STYLE: 1, SCRIPT: 1, HEAD: 1, TITLE: 1, META: 1, BLOCKQUOTE: 1, CODE: 1, PRE: 1 };
  var BLOCK_TAGS = { P: 1, DIV: 1, LI: 1, UL: 1, OL: 1, TR: 1, TD: 1, TH: 1, TABLE: 1, H1: 1, H2: 1, H3: 1, H4: 1, H5: 1, H6: 1, HR: 1, SECTION: 1 };

  /** Extrait le texte corrigeable du document HTML. */
  function extract(doc) {
    var text = "";
    var segments = [];
    var stopped = false;

    function walk(node) {
      if (stopped) return;
      if (node.nodeType === 3) {
        var v = node.nodeValue;
        if (v) {
          segments.push({ node: node, start: text.length, end: text.length + v.length });
          text += v;
        }
        return;
      }
      if (node.nodeType !== 1) return;
      var tag = node.tagName.toUpperCase();
      var id = node.id || "";
      var cls = typeof node.className === "string" ? node.className : "";
      if (THREAD_MARKERS.test(id)) { stopped = true; return; }
      if (SKIP_TAGS[tag] || SKIP_ID.test(id) || SKIP_CLASS.test(cls)) return;
      if (tag === "BR") { text += "\n"; return; }
      var block = BLOCK_TAGS[tag];
      if (block) text += "\n";
      for (var c = node.firstChild; c; c = c.nextSibling) walk(c);
      if (block) text += "\n";
    }

    walk(doc.body || doc.documentElement);
    return { text: text, segments: segments };
  }

  /** Découpe le texte en morceaux < MAX_CHARS, sur des retours à la ligne. */
  function chunks(text) {
    var out = [];
    var start = 0;
    while (start < text.length) {
      var end = Math.min(start + MAX_CHARS, text.length);
      if (end < text.length) {
        var nl = text.lastIndexOf("\n", end);
        if (nl > start) end = nl + 1;
      }
      out.push({ base: start, text: text.slice(start, end) });
      start = end;
    }
    return out;
  }

  /** Interroge LanguageTool et renvoie les fautes avec leur position dans le texte complet. */
  async function check(text, language) {
    var all = [];
    var parts = chunks(text);
    for (var i = 0; i < parts.length; i++) {
      if (!parts[i].text.trim()) continue;
      var body = new URLSearchParams();
      body.set("text", parts[i].text);
      body.set("language", language || "fr");
      var res = await fetch(API_URL, {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded", Accept: "application/json" },
        body: body.toString(),
      });
      if (res.status === 429) throw new Error("Trop de corrections en peu de temps : réessayez dans une minute.");
      if (!res.ok) throw new Error("Le service de correction a répondu une erreur (" + res.status + ").");
      var data = await res.json();
      (data.matches || []).forEach(function (m) {
        m.offset += parts[i].base;
        all.push(m);
      });
    }
    return all;
  }

  function isSentenceStart(text, offset) {
    for (var i = offset - 1; i >= 0; i--) {
      var ch = text[i];
      if (/\s/.test(ch)) { if (ch === "\n") return true; continue; }
      return /[.!?…:]/.test(ch);
    }
    return true;
  }

  function surroundingToken(text, offset, length) {
    var s = offset, e = offset + length;
    while (s > 0 && !/\s/.test(text[s - 1])) s--;
    while (e < text.length && !/\s/.test(text[e])) e++;
    return text.slice(s, e);
  }

  // Catégories LanguageTool corrigées : orthographe, grammaire (accords, conjugaison),
  // homonymes (a/à, ou/où, participe/infinitif). Le style et la typographie sont ignorés.
  var SPELLING_CATS = { TYPOS: 1 };
  var GRAMMAR_CATS = { GRAMMAR: 1, CAT_GRAMMAIRE: 1, CAT_HOMONYMES_PARONYMES: 1, CONFUSED_WORDS: 1, CAT_ELISION: 1 };

  function kind(m) {
    var rule = m.rule || {};
    var cat = (rule.category && rule.category.id) || "";
    if (rule.issueType === "misspelling" || SPELLING_CATS[cat]) return "misspelling";
    if (rule.issueType === "grammar" || GRAMMAR_CATS[cat]) return "grammar";
    return "";
  }

  /** Garde uniquement les fautes sûres à corriger automatiquement. */
  function filter(matches, text, dictionary) {
    var dict = {};
    (dictionary || []).forEach(function (w) { dict[w.toLowerCase()] = 1; });
    return matches.filter(function (m) {
      var type = kind(m);
      if (!type) return false;
      if (!m.replacements || !m.replacements.length) return false;
      var original = text.substr(m.offset, m.length);
      var replacement = m.replacements[0].value;
      if (!original.trim() || replacement === original) return false;
      if (dict[original.toLowerCase()]) return false;
      var token = surroundingToken(text, m.offset, m.length);
      if (/[@\/\\]|www\.|\.(fr|com|net|org)\b/i.test(token)) return false; // adresses, chemins, URL
      if (type === "misspelling") {
        if (/\d/.test(original)) return false;                                   // références, numéros
        if (original.length > 1 && original === original.toUpperCase()) return false; // sigles (ACS, VPN…)
        if (/^[A-ZÀ-Ý]/.test(original) && !isSentenceStart(text, m.offset)) return false; // noms propres
      }
      return true;
    }).map(function (m) {
      return {
        offset: m.offset,
        length: m.length,
        original: text.substr(m.offset, m.length),
        replacement: m.replacements[0].value,
        message: m.message || "",
        type: kind(m),
      };
    });
  }

  /** Applique les corrections dans les nœuds texte. Renvoie celles réellement appliquées. */
  function apply(extracted, fixes) {
    var bySeg = new Map();
    fixes.forEach(function (f) {
      var seg = extracted.segments.find(function (s) { return f.offset >= s.start && f.offset + f.length <= s.end; });
      if (!seg) return; // faute à cheval sur deux mises en forme : on n'y touche pas
      if (!bySeg.has(seg)) bySeg.set(seg, []);
      bySeg.get(seg).push(f);
    });
    var applied = [];
    bySeg.forEach(function (list, seg) {
      list.sort(function (a, b) { return b.offset - a.offset; });
      var v = seg.node.nodeValue;
      var lastStart = Infinity;
      list.forEach(function (f) {
        var s = f.offset - seg.start;
        if (s + f.length > lastStart) return; // chevauchement
        v = v.slice(0, s) + f.replacement + v.slice(s + f.length);
        lastStart = s;
        applied.push(f);
      });
      seg.node.nodeValue = v;
    });
    applied.sort(function (a, b) { return a.offset - b.offset; });
    return applied;
  }

  /**
   * Corrige un HTML complet. Renvoie { html, fixes }.
   * options : { language, dictionary }
   */
  async function correctHtml(html, options) {
    options = options || {};
    var doc = new DOMParser().parseFromString(html, "text/html");
    var ex = extract(doc);
    if (!ex.text.trim()) return { html: html, fixes: [] };
    var matches = await check(ex.text, options.language);
    var fixes = filter(matches, ex.text, options.dictionary);
    var applied = apply(ex, fixes);
    if (!applied.length) return { html: html, fixes: [] };
    var doctype = /^\s*<!doctype[^>]*>/i.exec(html);
    return { html: (doctype ? doctype[0] : "") + doc.documentElement.outerHTML, fixes: applied };
  }

  global.Correcteur = { correctHtml: correctHtml, _extract: extract, _filter: filter, _apply: apply };
})(window);
