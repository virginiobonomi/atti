/*
 * Atti — editor di clausole e minute dentro il pannello.
 * Il testo si scrive normalmente; le voci da compilare ("stop") sono riquadri: clic per modificarle.
 */
(function (root) {
  "use strict";

  var D = root.document;

  var TIPI = [
    { id: "testo", nome: "Testo libero" },
    { id: "scelta", nome: "Scelta tra valori" },
    { id: "data", nome: "Data" },
    { id: "numero", nome: "Numero" },
    { id: "importo", nome: "Importo" },
    { id: "comune", nome: "Comune" },
    { id: "cf", nome: "Codice fiscale / partita IVA" },
    { id: "elenco", nome: "Nome da elenco (parti)" },
    { id: "clausole", nome: "Lista di clausole" }
  ];
  var FORMATI_DATA = ["", "8/11/2026", "otto novembre duemilaventisei", "8 (otto) 11 (novembre) 2026 (duemilaventisei)",
    "8/11/26", "8 novembre 2026", "8 (otto) novembre 2026 (duemilaventisei)", "8 novembre 26", "8 (solo giorno)",
    "otto (giorno in lettere)", "11 (solo mese)", "novembre (mese in lettere)", "26 (anno a due cifre)", "2026 (solo anno)",
    "duemilaventisei (anno in lettere)", "otto novembre duemilaventisei (8.11.2026)"];
  var FORMATI_NUMERO = ["", "250000", "250.000", "250000 (duecentocinquantamila)", "250.000 (duecentocinquantamila)",
    "(duecentocinquantamila) 250000", "(duecentocinquantamila) 250.000", "duecentocinquantamila", "", "", "parte intera in lettere",
    "parte decimale in lettere", "250.000,50 (duecentocinquantamila euro e cinquanta centesimi)", "duecentocinquantamila (250.000)"];

  function el(tag, attr, figli) {
    var n = D.createElement(tag);
    Object.keys(attr || {}).forEach(function (k) {
      if (k === "testo") n.textContent = attr[k];
      else if (k === "classe") n.className = attr[k];
      else if (k.indexOf("on") === 0) n.addEventListener(k.slice(2), attr[k]);
      else n.setAttribute(k, attr[k]);
    });
    (figli || []).forEach(function (f) { if (f !== null && f !== undefined) n.appendChild(typeof f === "string" ? D.createTextNode(f) : f); });
    return n;
  }

  function tipoDi(st) {
    if (st.t === "LC") return "clausole";
    if (st.t === "D") return "data";
    if (st.t === "N") return st.k === "CU" ? "importo" : "numero";
    if (st.t === "CM") return "comune";
    if (st.t === "CF") return "cf";
    if (st.t === "VE") return "elenco";
    if (st.v && st.v.length) return "scelta";
    return "testo";
  }

  function descrivi(st) {
    var t = TIPI.filter(function (x) { return x.id === tipoDi(st); })[0];
    var s = t ? t.nome : "";
    if (st.t === "LC") s += st.f === 2 ? " (più clausole)" : " (una clausola)";
    return s;
  }

  // ---------------------------------------------------------------- modello -> DOM

  function chipStop(id, x, st) {
    var c = el("span", { classe: "ed-stop", contenteditable: "false", "data-id": id, title: (st ? descrivi(st) + " · " : "") + "clic per modificare" });
    c.textContent = (x.l || (st && st.n) || "voce").trim() || "voce";
    if (x.b) c.setAttribute("data-b", "1");
    if (x.i) c.setAttribute("data-i", "1");
    if (x.u) c.setAttribute("data-u", "1");
    // conserva gli spazi attorno all'etichetta originale
    var m = (x.l || "").match(/^(\s*)[\s\S]*?(\s*)$/);
    if (m) { c.setAttribute("data-pre", m[1]); c.setAttribute("data-post", m[2]); }
    return c;
  }

  function versoDom(doc, cont) {
    cont.innerHTML = "";
    doc.p.forEach(function (p) {
      var par = el("p", { classe: "ed-p" });
      if (p.a && p.a !== "l") par.setAttribute("data-a", p.a);
      if (p.li) par.setAttribute("data-li", p.li);
      if (p.fi) par.setAttribute("data-fi", p.fi);
      applicaStileParagrafo(par);
      p.x.forEach(function (x) {
        if (x.br) { par.appendChild(el("br")); return; }
        if (x.s !== undefined) {
          par.appendChild(chipStop(x.s, x, doc.s[x.s]));   // gli spazi attorno all'etichetta restano nei data-pre/post
          return;
        }
        var n = D.createTextNode(x.t);
        if (x.k) n = el("span", { classe: "ed-sc" }, [n]);
        if (x.u) n = el("u", {}, [n]);
        if (x.i) n = el("i", {}, [n]);
        if (x.b) n = el("b", {}, [n]);
        par.appendChild(n);
      });
      // un a capo finale si vede solo con un secondo <br> (che in lettura viene tolto)
      if (!par.childNodes.length || (p.x.length && p.x[p.x.length - 1].br)) par.appendChild(el("br"));
      cont.appendChild(par);
    });
    if (!doc.p.length) cont.appendChild(el("p", { classe: "ed-p" }, [el("br")]));
  }

  function applicaStileParagrafo(par) {
    var a = par.getAttribute("data-a");
    par.style.textAlign = a === "c" ? "center" : a === "r" ? "right" : a === "j" ? "justify" : "";
    var li = +par.getAttribute("data-li") || 0;
    par.style.marginLeft = li ? (li / 20) + "pt" : "";
  }

  // ---------------------------------------------------------------- DOM -> modello

  function daDom(cont, defs) {
    var paragrafi = [];
    var blocchi = [];
    var sciolti = [];
    Array.prototype.forEach.call(cont.childNodes, function (n) {
      if (n.nodeType === 1 && /^(P|DIV|H\d|LI|BLOCKQUOTE)$/.test(n.tagName)) {
        if (sciolti.length) { blocchi.push({ nodi: sciolti }); sciolti = []; }
        blocchi.push({ el: n, nodi: Array.prototype.slice.call(n.childNodes) });
      } else sciolti.push(n);
    });
    if (sciolti.length) blocchi.push({ nodi: sciolti });
    var usati = {};
    blocchi.forEach(function (b) {
      var p = { x: [] };
      if (b.el) {
        var a = b.el.getAttribute("data-a") || ({ center: "c", right: "r", justify: "j" })[b.el.style.textAlign] || "";
        if (a && a !== "l") p.a = a;
        if (b.el.getAttribute("data-li")) p.li = +b.el.getAttribute("data-li");
        if (b.el.getAttribute("data-fi")) p.fi = +b.el.getAttribute("data-fi");
      }
      function aggiungiTesto(t, f) {
        if (!t) return;
        var ult = p.x[p.x.length - 1];
        if (ult && ult.t !== undefined && !!ult.b === !!f.b && !!ult.i === !!f.i && !!ult.u === !!f.u && !!ult.k === !!f.k) { ult.t += t; return; }
        var x = { t: t };
        ["b", "i", "u", "k"].forEach(function (k) { if (f[k]) x[k] = 1; });
        p.x.push(x);
      }
      function giro(n, f) {
        if (n.nodeType === 3) { aggiungiTesto(n.nodeValue.replace(/​/g, ""), f); return; }
        if (n.nodeType !== 1) return;
        if (n.classList.contains("ed-stop")) {
          var id = n.getAttribute("data-id");
          var x = { s: id, l: (n.getAttribute("data-pre") || "") + n.textContent + (n.getAttribute("data-post") || "") };
          if (n.getAttribute("data-b")) x.b = 1;
          if (n.getAttribute("data-i")) x.i = 1;
          if (n.getAttribute("data-u")) x.u = 1;
          p.x.push(x);
          usati[id] = 1;
          return;
        }
        if (n.tagName === "BR") { p.x.push({ br: 1 }); return; }
        if (/^(P|DIV)$/.test(n.tagName)) {   // paragrafo annidato creato dall'editor del browser
          n.childNodes.forEach(function (c) { giro(c, f); });
          return;
        }
        var g = { b: f.b, i: f.i, u: f.u, k: f.k };
        var stile = n.style || {};
        if (/^(B|STRONG)$/.test(n.tagName) || stile.fontWeight === "bold" || +stile.fontWeight >= 600) g.b = 1;
        if (/^(I|EM)$/.test(n.tagName) || stile.fontStyle === "italic") g.i = 1;
        if (n.tagName === "U" || /underline/.test(stile.textDecoration || "")) g.u = 1;
        if (n.classList.contains("ed-sc")) g.k = 1;
        Array.prototype.forEach.call(n.childNodes, function (c) { giro(c, g); });
      }
      b.nodi.forEach(function (n) { giro(n, {}); });
      // il <br> finale che il browser mette nei paragrafi vuoti non è un a capo
      if (p.x.length && p.x[p.x.length - 1].br) p.x.pop();
      paragrafi.push(p);
    });
    // niente paragrafi vuoti in coda
    while (paragrafi.length && !paragrafi[paragrafi.length - 1].x.length) paragrafi.pop();
    var s = {};
    Object.keys(defs).forEach(function (id) { if (usati[id]) s[id] = defs[id]; });
    return { p: paragrafi, s: s };
  }

  // ---------------------------------------------------------------- scheda della voce

  function schedaStop(st, opz, fatto) {
    var box = el("div", { classe: "ed-scheda" });
    var tipo = tipoDi(st);
    var nome = el("input", { type: "text", value: st.n || "" });
    var selTipo = el("select");
    TIPI.forEach(function (t) { var o = el("option", { value: t.id, testo: t.nome }); if (t.id === tipo) o.selected = true; selTipo.appendChild(o); });
    var corpo = el("div");
    var stato = { v: (st.v || []).slice(), f: st.f, ft: st.ft || "", pr: st.pr || "", po: st.po || "", ve: !!st.ve };

    function campo(etichetta, controllo, nota) {
      return el("label", { classe: "ed-campo" }, [el("span", { testo: etichetta }), controllo, nota ? el("small", { testo: nota }) : null]);
    }

    function disegna() {
      corpo.innerHTML = "";
      var t = selTipo.value;
      if (t === "scelta") {
        var ta = el("textarea", { rows: "4" });
        ta.value = (st.t === "T" ? stato.v : []).join("\n");
        ta.addEventListener("input", function () { stato.v = ta.value.split("\n"); });
        corpo.appendChild(campo("Valori possibili, uno per riga (una riga vuota = nessun testo)", ta));
      }
      if (t === "data" || t === "numero" || t === "importo") {
        var lista = t === "data" ? FORMATI_DATA : FORMATI_NUMERO;
        var sel = el("select");
        lista.forEach(function (d, i) { if (i && d) { var o = el("option", { value: String(i), testo: d }); if (i === (stato.f || (t === "data" ? 5 : 4))) o.selected = true; sel.appendChild(o); } });
        stato.f = +sel.value;
        sel.addEventListener("change", function () { stato.f = +sel.value; });
        corpo.appendChild(campo("Come si scrive nell'atto", sel));
      }
      if (t === "clausole") {
        var multi = el("select");
        [["1", "Una sola clausola"], ["2", "Più clausole (anche ripetute)"]].forEach(function (o) {
          var op = el("option", { value: o[0], testo: o[1] }); if (+o[0] === (st.t === "LC" ? st.f : 1)) op.selected = true; multi.appendChild(op);
        });
        stato.f = +multi.value;
        multi.addEventListener("change", function () { stato.f = +multi.value; });
        corpo.appendChild(campo("Scelta", multi));
        if (st.t !== "LC") stato.v = [];
        var elenco = el("div", { classe: "ed-fonti" });
        var ridisegna = function () {
          elenco.innerHTML = "";
          stato.v.forEach(function (c, i) {
            var testo = c[1] === "*" ? "Tutta la categoria «" + c[0] + "»" : (opz.nomeClausola(c[1]) || c[1]);
            elenco.appendChild(el("div", { classe: "scelta-fatta" }, [el("span", { testo: testo }),
              el("button", { type: "button", testo: "✕", onclick: function () { stato.v.splice(i, 1); ridisegna(); } })]));
          });
          if (!stato.v.length) elenco.appendChild(el("small", { testo: "Nessuna clausola: aggiungi una categoria o singole clausole." }));
        };
        ridisegna();
        corpo.appendChild(campo("Clausole proposte", elenco));
        var cat = el("input", { type: "text", list: "ed-categorie", placeholder: "categoria (es. Provenienza)" });
        corpo.appendChild(el("div", { classe: "ed-riga" }, [cat, el("button", { type: "button", testo: "Aggiungi categoria", onclick: function () {
          var c = cat.value.trim().toLowerCase(); if (!c) return; stato.v.push([c, "*"]); cat.value = ""; ridisegna();
        } })]));
        var cl = el("input", { type: "text", list: "ed-clausole", placeholder: "singola clausola (scrivi per cercare)" });
        corpo.appendChild(el("div", { classe: "ed-riga" }, [cl, el("button", { type: "button", testo: "Aggiungi clausola", onclick: function () {
          var k = opz.chiaveClausola(cl.value.trim()); if (!k) { cl.setCustomValidity("Clausola non trovata"); cl.reportValidity(); return; }
          cl.setCustomValidity(""); stato.v.push([k.split("/").slice(0, -1).join("/"), k]); cl.value = ""; ridisegna();
        } })]));
      }
      if (t === "testo" || t === "scelta" || t === "comune" || t === "elenco") {
        var ft = el("select");
        [["", "Come scritto"], ["Minuscolo", "minuscolo"], ["MAIUSCOLO", "MAIUSCOLO"], ["InizialiMaiuscole", "Iniziali Maiuscole"]].forEach(function (o) {
          var op = el("option", { value: o[0], testo: o[1] }); if (o[0] === stato.ft) op.selected = true; ft.appendChild(op);
        });
        ft.addEventListener("change", function () { stato.ft = ft.value; });
        corpo.appendChild(campo("Maiuscole e minuscole", ft));
      }
      if (t !== "clausole") {
        var pr = el("input", { type: "text", value: stato.pr }), po = el("input", { type: "text", value: stato.po });
        pr.addEventListener("input", function () { stato.pr = pr.value; });
        po.addEventListener("input", function () { stato.po = po.value; });
        corpo.appendChild(el("div", { classe: "ed-riga" }, [campo("Testo prima (se compilata)", pr), campo("Testo dopo", po)]));
      }
      if (t === "testo" || t === "scelta") {
        var cb = el("input", { type: "checkbox" }); cb.checked = stato.ve;
        cb.addEventListener("change", function () { stato.ve = cb.checked; });
        corpo.appendChild(el("label", { classe: "check" }, [cb, "Proponi il valore nei campi \"nome\" (es. nome delle parti)"]));
      }
    }
    selTipo.addEventListener("change", disegna);
    disegna();

    box.appendChild(el("div", { classe: "ed-scheda-testa", testo: "Voce da compilare" }));
    box.appendChild(campo("Nome (è anche il testo mostrato finché la voce non è compilata)", nome));
    box.appendChild(campo("Tipo", selTipo));
    box.appendChild(corpo);
    box.appendChild(el("div", { classe: "bottoni" }, [
      el("button", { type: "button", classe: "primario", testo: "Applica", onclick: function () {
        var t = selTipo.value, n = nome.value.trim() || "voce";
        var nuovo = { n: n };
        if (t === "testo" || t === "scelta") { nuovo.t = "T"; nuovo.f = 1; nuovo.k = "TE"; }
        if (t === "scelta") nuovo.v = stato.v.length ? stato.v : [""];
        if (t === "data") { nuovo.t = "D"; nuovo.f = stato.f || 5; nuovo.k = "DT"; }
        if (t === "numero") { nuovo.t = "N"; nuovo.f = stato.f || 4; nuovo.k = "NR"; }
        if (t === "importo") { nuovo.t = "N"; nuovo.f = stato.f || 4; nuovo.k = "CU"; }
        if (t === "comune") { nuovo.t = "CM"; nuovo.f = 1; nuovo.k = "CM"; }
        if (t === "cf") { nuovo.t = "CF"; nuovo.f = 2; nuovo.k = "CF"; }
        if (t === "elenco") { nuovo.t = "VE"; nuovo.f = 1; nuovo.k = "VE"; }
        if (t === "clausole") { nuovo.t = "LC"; nuovo.f = stato.f || 1; nuovo.k = "LC"; nuovo.v = stato.v; }
        if (stato.ft && t !== "clausole") nuovo.ft = stato.ft;
        if (stato.pr && t !== "clausole") nuovo.pr = stato.pr;
        if (stato.po && t !== "clausole") nuovo.po = stato.po;
        if (stato.ve && (t === "testo" || t === "scelta")) nuovo.ve = 1;
        if (st.co) nuovo.co = st.co;
        fatto(nuovo);
      } }),
      el("button", { type: "button", testo: "Togli la voce dal testo", onclick: function () { fatto(null); } }),
      el("button", { type: "button", classe: "link", testo: "Chiudi", onclick: function () { fatto(undefined); } })
    ]));
    return box;
  }

  // ---------------------------------------------------------------- editor

  /*
   * opz: { sez, chiave, doc, nuovo, categorie: [nomi], clausole: [{chiave, nome}], info,
   *        salva(doc) -> Promise, elimina() -> Promise, ripristina() -> Promise, chiudi() }
   */
  function apri(contenitore, opz) {
    var doc = JSON.parse(JSON.stringify(opz.doc || { nome: "", cat: "", p: [], s: {} }));
    var defs = JSON.parse(JSON.stringify(doc.s || {}));
    var cont = contenitore;
    cont.innerHTML = "";
    cont.hidden = false;

    var mappaNomi = {}, mappaChiavi = {};
    (opz.clausole || []).forEach(function (c) { mappaNomi[c.chiave] = c.nome + " — " + c.cat; mappaChiavi[(c.nome + " — " + c.cat).toLowerCase()] = c.chiave; mappaChiavi[c.nome.toLowerCase()] = mappaChiavi[c.nome.toLowerCase()] || c.chiave; });
    var dlCat = el("datalist", { id: "ed-categorie" });
    (opz.categorie || []).forEach(function (c) { dlCat.appendChild(el("option", { value: c })); });
    var dlCl = el("datalist", { id: "ed-clausole" });
    (opz.clausole || []).forEach(function (c) { dlCl.appendChild(el("option", { value: c.nome + " — " + c.cat })); });

    var titolo = opz.sez === "minute" ? (opz.nuovo ? "Nuova minuta" : "Modifica minuta") : (opz.nuovo ? "Nuova clausola" : "Modifica clausola");
    var nome = el("input", { type: "text", value: doc.nome || "", placeholder: "nome" });
    var cat = el("input", { type: "text", value: doc.cat || "", list: "ed-categorie", placeholder: opz.sez === "minute" ? "cartella" : "categoria" });
    var stato = el("p", { classe: "stato" });
    var area = el("div", { classe: "ed-testo", contenteditable: "true", spellcheck: "true" });
    var scheda = el("div", { classe: "ed-scheda-cont" });

    function comando(c, v) { area.focus(); try { D.execCommand(c, false, v); } catch (e) { /* ignora */ } }
    function paragrafoCorrente() {
      var s = root.getSelection && root.getSelection();
      var n = s && s.anchorNode;
      while (n && n !== area && !(n.nodeType === 1 && /^(P|DIV)$/.test(n.tagName) && n.parentNode === area)) n = n.parentNode;
      return n && n !== area ? n : null;
    }
    function allinea(a) {
      var p = paragrafoCorrente();
      if (!p) return;
      if (a === "l") p.removeAttribute("data-a"); else p.setAttribute("data-a", a);
      applicaStileParagrafo(p);
    }
    function nuovoId() {
      var max = 0;
      Object.keys(defs).forEach(function (k) { if (+k > max) max = +k; });
      Array.prototype.forEach.call(area.querySelectorAll(".ed-stop"), function (c) { if (+c.getAttribute("data-id") > max) max = +c.getAttribute("data-id"); });
      return String(max + 1);
    }
    function nuovaVoce() {
      area.focus();
      var s = root.getSelection();
      var testo = s && s.rangeCount ? s.toString().trim() : "";
      var id = nuovoId();
      defs[id] = { n: testo || "nuova voce", t: "T", f: 1, k: "TE" };
      var chip = chipStop(id, { l: defs[id].n }, defs[id]);
      if (s && s.rangeCount && area.contains(s.anchorNode)) {
        var r = s.getRangeAt(0);
        r.deleteContents();
        r.insertNode(chip);
        r.setStartAfter(chip); r.collapse(true);
        s.removeAllRanges(); s.addRange(r);
      } else {
        var ult = area.lastElementChild || area.appendChild(el("p", { classe: "ed-p" }));
        ult.appendChild(chip);
      }
      apriScheda(chip);
    }
    function apriScheda(chip) {
      var id = chip.getAttribute("data-id");
      Array.prototype.forEach.call(area.querySelectorAll(".ed-stop.sel"), function (c) { c.classList.remove("sel"); });
      chip.classList.add("sel");
      scheda.innerHTML = "";
      scheda.appendChild(schedaStop(defs[id] || { n: chip.textContent, t: "T", f: 1 }, {
        nomeClausola: function (k) { return mappaNomi[k]; },
        chiaveClausola: function (testo) { return mappaChiavi[testo.toLowerCase()]; }
      }, function (nuovo) {
        if (nuovo === null) chip.parentNode.removeChild(chip);
        else if (nuovo) { defs[id] = nuovo; chip.textContent = nuovo.n; chip.title = descrivi(nuovo) + " · clic per modificare"; }
        chip.classList.remove("sel");
        scheda.innerHTML = "";
      }));
      scheda.scrollIntoView({ block: "nearest" });
    }
    area.addEventListener("click", function (e) {
      var c = e.target.closest && e.target.closest(".ed-stop");
      if (c) { e.preventDefault(); apriScheda(c); }
    });
    // incolla solo testo semplice
    area.addEventListener("paste", function (e) {
      var t = e.clipboardData && e.clipboardData.getData("text/plain");
      if (t === undefined || t === null) return;
      e.preventDefault();
      D.execCommand("insertText", false, t);
    });

    function stb(testo, titoloBtn, f, classe) { return el("button", { type: "button", classe: "ed-btn" + (classe ? " " + classe : ""), title: titoloBtn, testo: testo, onmousedown: function (e) { e.preventDefault(); }, onclick: f }); }
    var barra = el("div", { classe: "ed-barra" }, [
      stb("G", "Grassetto", function () { comando("bold"); }, "b"),
      stb("C", "Corsivo", function () { comando("italic"); }, "i"),
      stb("S", "Sottolineato", function () { comando("underline"); }, "u"),
      stb("⇤", "Allinea a sinistra", function () { allinea("l"); }),
      stb("↔", "Centra", function () { allinea("c"); }),
      stb("⇥", "Allinea a destra", function () { allinea("r"); }),
      stb("☰", "Giustifica", function () { allinea("j"); }),
      stb("＋ Voce", "Nuova voce da compilare (seleziona prima il testo che diventa il suo nome)", nuovaVoce, "voce")
    ]);

    var info = opz.info ? el("p", { classe: "nota", testo: (opz.info.eliminata ? "Eliminata" : "Modificata") + (opz.info.autore ? " da " + opz.info.autore : "") +
      (opz.info.data ? " il " + new Date(opz.info.data).toLocaleString("it-IT") : "") + ". Le versioni precedenti restano nella cronologia del file su SharePoint." }) : null;

    function occupato(s) { Array.prototype.forEach.call(cont.querySelectorAll("button"), function (b) { b.disabled = s; }); }
    function esegui(promessa, msg) {
      occupato(true); stato.textContent = "Salvataggio…"; stato.className = "stato";
      promessa.then(function () { stato.textContent = msg; stato.className = "stato ok"; occupato(false); })
        .catch(function (e) { stato.textContent = e.message || String(e); stato.className = "stato err"; occupato(false); });
    }

    var bottoni = el("div", { classe: "bottoni" }, [
      el("button", { type: "button", classe: "primario", testo: "Salva nel formulario", onclick: function () {
        if (!nome.value.trim()) { stato.textContent = "Manca il nome."; stato.className = "stato err"; return; }
        var nuovo = daDom(area, defs);
        nuovo.nome = nome.value.trim();
        nuovo.cat = cat.value.trim();
        esegui(opz.salva(nuovo), "Salvato: da ora lo vedono tutti.");
      } }),
      el("button", { type: "button", testo: "Chiudi", onclick: function () { cont.hidden = true; cont.innerHTML = ""; if (opz.chiudi) opz.chiudi(); } })
    ]);
    var altri = el("div", { classe: "bottoni" });
    if (!opz.nuovo && opz.elimina) altri.appendChild(el("button", { type: "button", classe: "link", testo: opz.sez === "minute" ? "Elimina la minuta" : "Elimina la clausola", onclick: function () {
      if (altri.getAttribute("data-conferma") !== "1") { altri.setAttribute("data-conferma", "1"); stato.textContent = "Clicca di nuovo per confermare l'eliminazione."; stato.className = "stato err"; return; }
      esegui(opz.elimina(), "Eliminata.");
    } }));
    if (opz.info && opz.ripristina) altri.appendChild(el("button", { type: "button", classe: "link", testo: "Torna al testo originale", onclick: function () {
      esegui(opz.ripristina(), "Ripristinato il testo originale.");
    } }));

    cont.appendChild(el("div", { classe: "velo-testa" }, [el("strong", { testo: titolo }), el("button", { classe: "chiudi", type: "button", testo: "✕", onclick: function () { cont.hidden = true; cont.innerHTML = ""; if (opz.chiudi) opz.chiudi(); } })]));
    cont.appendChild(el("div", { classe: "ed-corpo" }, [
      dlCat, dlCl,
      el("div", { classe: "ed-riga" }, [nome, cat]),
      info,
      barra, area, scheda,
      el("p", { classe: "nota", testo: "Scrivi il testo normalmente. I riquadri azzurri sono le voci da compilare: clic su un riquadro per cambiarne tipo e valori, \"＋ Voce\" per crearne una nuova." }),
      bottoni, altri, stato
    ]));
    versoDom(doc, area);
    return { leggi: function () { return daDom(area, defs); } };
  }

  root.Editor = { apri: apri, versoDom: versoDom, daDom: daDom, descrivi: descrivi };
})(this);
