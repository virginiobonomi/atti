/* Pannello Atti: sceglie la minuta, raccoglie scelte e dati, scrive l'atto in Word. */
(function () {
  "use strict";

  var $ = function (id) { return document.getElementById(id); };
  var inWord = false;
  var F = null, M = null;
  var connesso = false, infoMod = {};   // collegato a SharePoint; modifiche al formulario di base
  var minuta = null, R = {};
  var ui = { filtro: "", soloDaFare: false, anteprime: {}, espansi: {} };
  var CHIAVE_OPZ = "atti.opzioni";
  var opz = { formato: true, segnaposto: true };
  try { var o = JSON.parse(window.localStorage.getItem(CHIAVE_OPZ) || "null"); if (o) { opz.formato = o.formato !== false; opz.segnaposto = o.segnaposto !== false; } } catch (e) { /* ignora */ }

  // ---------------------------------------------------------------- formulario (IndexedDB)

  function idb(scrivi, valore) {
    return new Promise(function (ok, ko) {
      if (!window.indexedDB) { ko(new Error("IndexedDB non disponibile")); return; }
      var rq = window.indexedDB.open("atti", 1);
      rq.onupgradeneeded = function () { rq.result.createObjectStore("dati"); };
      rq.onerror = function () { ko(rq.error); };
      rq.onsuccess = function () {
        var tx = rq.result.transaction("dati", scrivi ? "readwrite" : "readonly");
        var st = tx.objectStore("dati");
        var r = scrivi ? st.put(valore, "formulario") : st.get("formulario");
        r.onsuccess = function () { ok(r.result); };
        r.onerror = function () { ko(r.error); };
      };
    });
  }

  function usaFormulario(testo, salvare) {
    var dati;
    try { dati = JSON.parse(testo); } catch (e) { throw new Error("Il file non è un formulario valido."); }
    if (!dati || !dati.minute || !dati.clausole) throw new Error("Il file non è un formulario valido.");
    F = dati; M = Atti.crea(F);
    $("versione").textContent = "Formulario del " + (F.versione || "?") + ": " + Object.keys(F.minute).length +
      " minute, " + Object.keys(F.clausole).length + " clausole.";
    return salvare ? idb(true, testo).catch(function () { /* resta in memoria solo per questa sessione */ }) : Promise.resolve();
  }

  function costruisci(esito, online) {
    var dati = JSON.parse(esito.base);
    infoMod = Archivio.applica(dati, esito.modifiche);
    F = dati; M = Atti.crea(F);
    connesso = online;
    aggiornaVersione(online);
  }

  function aggiornaVersione(online) {
    var n = Object.keys(infoMod).length;
    $("versione").textContent = (online ? "Formulario dello studio su SharePoint" : "Copia locale del formulario (SharePoint non raggiungibile)") +
      ": " + Object.keys(F.minute).length + " minute, " + Object.keys(F.clausole).length + " clausole" +
      (n ? ", " + n + (n === 1 ? " modifica" : " modifiche") + " dello studio" : "") + ".";
  }

  function daSharePoint(interattivo) {
    if (!window.Archivio) return Promise.reject(new Error("Modulo SharePoint non caricato."));
    $("infoFormulario").textContent = "Collegamento al formulario dello studio…";
    return (interattivo ? Archivio.accedi() : Promise.resolve()).then(function () {
      return Archivio.sincronizza(function (msg) { $("infoFormulario").textContent = msg; });
    }).then(function (esito) {
      costruisci(esito, true);
      $("infoFormulario").textContent = "Minute e clausole dello studio, con le scelte e i dati da compilare.";
    });
  }

  function avviaFormulario() {
    return daSharePoint(false).catch(function (e) {
      if (e && e.accesso) { $("secAccesso").hidden = false; $("infoFormulario").textContent = "Serve l'accesso con l'account dello studio."; return; }
      // SharePoint non raggiungibile: ultima copia scaricata, oppure formulario caricato da file
      return (window.Archivio ? Archivio.daCache() : Promise.resolve(null)).then(function (c) {
        if (c) { costruisci(c, false); stato("SharePoint non raggiungibile (" + (e && e.message ? e.message : e) + "): uso l'ultima copia scaricata.", "err"); return; }
        return idb(false).then(function (testo) {
          if (testo) return usaFormulario(testo, false);
          stato("SharePoint non raggiungibile" + (e && e.message ? ": " + e.message : "") + ".", "err");
        }).catch(function () {});
      }).then(function () { $("infoFormulario").textContent = "Minute e clausole dello studio, con le scelte e i dati da compilare."; });
    });
  }

  function caricaFile(file) {
    if (!file) return;
    stato("Lettura del formulario…");
    file.text().then(function (t) { return usaFormulario(t, true); }).then(function () {
      stato("Formulario caricato.", "ok");
      ripristinaDaDocumento();
      disegnaTutto();
    }).catch(function (e) { stato(e.message || String(e), "err"); });
  }

  // ---------------------------------------------------------------- stato dell'atto (nel documento)

  var timerSalva = null;
  function salvaStato() {
    clearTimeout(timerSalva);
    timerSalva = setTimeout(function () {
      var dati = { m: minuta, r: R };
      if (inWord && Office.context.document.settings) {
        try {
          Office.context.document.settings.set("atti", dati);
          Office.context.document.settings.saveAsync(function () { /* salvato con il documento */ });
        } catch (e) { /* ignora */ }
      } else {
        try { window.localStorage.setItem("atti.prova", JSON.stringify(dati)); } catch (e) { /* ignora */ }
      }
    }, 500);
  }

  function ripristinaDaDocumento() {
    var dati = null;
    try {
      if (inWord && Office.context.document.settings) dati = Office.context.document.settings.get("atti");
      else dati = JSON.parse(window.localStorage.getItem("atti.prova") || "null");
    } catch (e) { dati = null; }
    if (dati && dati.m && F && F.minute[dati.m]) { minuta = dati.m; R = dati.r || {}; }
  }

  // ---------------------------------------------------------------- utilità interfaccia

  function el(tag, attr, figli) {
    var n = document.createElement(tag);
    Object.keys(attr || {}).forEach(function (k) {
      if (k === "testo") n.textContent = attr[k];
      else if (k === "classe") n.className = attr[k];
      else if (k.indexOf("on") === 0) n.addEventListener(k.slice(2), attr[k]);
      else n.setAttribute(k, attr[k]);
    });
    (figli || []).forEach(function (f) { if (f) n.appendChild(typeof f === "string" ? document.createTextNode(f) : f); });
    return n;
  }

  function stato(msg, tipo) {
    var s = $("stato");
    s.textContent = msg || "";
    s.className = "stato" + (tipo ? " " + tipo : "");
  }

  function testoClausola(d) {
    return d.p.map(function (p) {
      return p.x.map(function (x) { return x.t !== undefined ? x.t : (x.s !== undefined ? "[" + x.l.trim() + "]" : "\n"); }).join("");
    }).join("\n");
  }

  // ---------------------------------------------------------------- scelta della minuta

  function disegnaMinute() {
    var cont = $("listaMinute");
    cont.innerHTML = "";
    var parole = $("cercaMinuta").value.toLowerCase().split(/\s+/).filter(Boolean);
    var lista = M.minute().filter(function (m) {
      var t = (m.nome + " " + m.cat).toLowerCase();
      return parole.every(function (p) { return t.indexOf(p) >= 0; });
    });
    // prima le minute dello studio
    lista.sort(function (a, b) {
      var sa = /varese notai/i.test(a.cat) ? 0 : 1, sb = /varese notai/i.test(b.cat) ? 0 : 1;
      return sa - sb || a.cat.localeCompare(b.cat) || a.nome.localeCompare(b.nome);
    });
    var gruppo = null, n = 0;
    lista.forEach(function (m) {
      if (++n > 300) return;
      if (m.cat !== gruppo) { gruppo = m.cat; cont.appendChild(el("div", { classe: "gruppo", testo: m.cat || "Minute" })); }
      cont.appendChild(el("button", { classe: "minuta", type: "button", testo: m.nome, onclick: function () { scegliMinuta(m.chiave); } }));
    });
    if (!lista.length) cont.appendChild(el("div", { classe: "vuota", testo: "Nessuna minuta trovata." }));
  }

  function scegliMinuta(k) {
    if (minuta && minuta !== k && Object.keys(R).length) R = {};
    minuta = k;
    ui.espansi = {}; ui.anteprime = {};
    salvaStato();
    disegnaTutto();
    window.scrollTo(0, 0);
  }

  // ---------------------------------------------------------------- risposte

  function imposta(chiave, valore) {
    if (valore === undefined) delete R[chiave]; else R[chiave] = valore;
    salvaStato();
  }

  // quando si toglie o si sposta una clausola da una lista, le risposte annidate seguono
  function rinumera(percorso, mappa) {
    var pref = percorso + "[";
    var nuove = {};
    Object.keys(R).forEach(function (k) {
      if (k.indexOf(pref) !== 0) { nuove[k] = R[k]; return; }
      var m = k.slice(pref.length).match(/^(\d+)\](.*)$/);
      if (!m) { nuove[k] = R[k]; return; }
      var nuovo = mappa(+m[1]);
      if (nuovo === null || nuovo === undefined) return;
      nuove[pref + nuovo + "]" + m[2]] = R[k];
    });
    R = nuove;
  }

  // ---------------------------------------------------------------- elenco degli stop

  function nomiSuggeriti(elenco) {
    var visti = {}, out = [];
    elenco.forEach(function (v) {
      var st = v.stop;
      var adatto = st.ve || st.t === "VE" || /cognome|nome|denominazione|ragione sociale/i.test(st.n || "");
      if (!adatto || typeof v.valore !== "string") return;
      v.valore.split(/\s*,\s*|\s+e\s+/).forEach(function (s) {
        s = s.trim();
        if (s && !visti[s]) { visti[s] = 1; out.push(s); }
      });
    });
    return out;
  }

  function disegnaStop() {
    var cont = $("stop");
    cont.innerHTML = "";
    var elenco = M.elenco(minuta, R);
    var dl = el("datalist", { id: "nomi" });
    nomiSuggeriti(elenco).forEach(function (s) { dl.appendChild(el("option", { value: s })); });
    cont.appendChild(dl);
    var parole = ui.filtro.toLowerCase().split(/\s+/).filter(Boolean);
    var sez = null, mostrati = 0;
    elenco.forEach(function (v) {
      if (v.duplicato) return;
      if (ui.soloDaFare && v.compilato && v.stop.t !== "LC") return;
      if (ui.soloDaFare && v.compilato && v.stop.t === "LC" && !ui.espansi[v.percorso]) return;
      if (parole.length) {
        var t = (v.etichetta + " " + v.prima + " " + v.dopo + " " + v.clausola + " " + v.sezione).toLowerCase();
        if (!parole.every(function (p) { return t.indexOf(p) >= 0; })) return;
      }
      if (v.livello === 0 && v.sezione !== sez) {
        sez = v.sezione;
        if (sez) cont.appendChild(el("div", { classe: "sezione", testo: sez }));
      }
      cont.appendChild(voce(v));
      mostrati++;
    });
    if (!mostrati) cont.appendChild(el("div", { classe: "vuota", testo: ui.soloDaFare ? "Tutte le voci sono compilate." : "Nessuna voce." }));
    aggiornaConteggio(elenco);
  }

  function aggiornaConteggio(elenco) {
    elenco = elenco || M.elenco(minuta, R);
    var daFare = elenco.filter(function (v) { return !v.duplicato && !v.compilato; }).length;
    $("conteggio").textContent = daFare === 0 ? "Tutte le voci sono compilate." :
      (daFare === 1 ? "1 voce da compilare" : daFare + " voci da compilare");
  }

  function voce(v) {
    var st = v.stop;
    var classi = "voce" + (v.compilato ? " fatta" : "") + (v.livello ? " annidata" + (v.livello > 1 ? " annidata" + Math.min(v.livello, 3) : "") : "");
    var box = el("div", { classe: classi });
    if (v.clausola) box.appendChild(el("div", { classe: "clausola-di", testo: "↳ " + v.clausola }));
    var azioni = el("div", { classe: "voce-azioni" });
    box.appendChild(el("div", { classe: "voce-testa" }, [el("span", { classe: "voce-nome", testo: v.etichetta }), azioni]));
    if (v.prima.trim() || v.dopo.trim()) {
      box.appendChild(el("div", { classe: "contesto" }, [
        "…" + v.prima.replace(/^\s+/, ""), el("mark", { testo: "▢" }), v.dopo + "…"
      ]));
    }
    if (st.t === "LC") widgetClausole(v, box, azioni);
    else widgetValore(v, box, azioni);
    return box;
  }

  function widgetValore(v, box, azioni) {
    var st = v.stop, chiave = v.chiaveValore;
    var valore = v.valore === undefined || v.valore === null ? "" : String(v.valore);
    var resa = el("div", { classe: "resa" });
    var input;
    function aggiornaResa() {
      var r = R[chiave];
      if (r === undefined) { resa.textContent = ""; return; }
      var f = st.t === "D" && v.condiviso && r ? Atti.formattaData(r, 2) : Atti.formatta(st, r);
      resa.textContent = f === "" ? "(omesso)" : (f !== r ? "→ " + f : "");
      if (v.condiviso && f) resa.textContent += " (vale per tutto l'atto)";
    }
    function cambia(val, daInput) {
      imposta(chiave, val);
      box.classList.toggle("fatta", R[chiave] !== undefined);
      if (!daInput && input) input.value = val || "";
      aggiornaResa();
      segnaScelte();
      aggiornaConteggio();
    }
    var scelte = null;
    function segnaScelte() {
      if (!scelte) return;
      Array.prototype.forEach.call(scelte.children, function (b) { b.classList.toggle("attiva", R[chiave] === b.getAttribute("data-v")); });
    }
    if (st.v && st.v.length) {
      scelte = el("div", { classe: "scelte" });
      st.v.forEach(function (opzione) {
        scelte.appendChild(el("button", { classe: "scelta", type: "button", "data-v": opzione, testo: opzione.trim() === "" ? "(niente)" : opzione,
          onclick: function () { cambia(opzione); } }));
      });
      box.appendChild(scelte);
    }
    if (st.t === "D") {
      input = el("input", { type: "date" });
      var d = valore.match(/^\d{4}-\d{2}-\d{2}$/) ? valore : "";
      input.value = d;
    } else {
      input = el("input", { type: "text", placeholder: st.t === "N" ? "numero (es. 250000 o 1.500,50)" : (st.t === "CF" ? "codice fiscale" : "") });
      input.value = valore;
      if (st.t === "VE" || st.ve || /cognome|nome/i.test(st.n || "")) input.setAttribute("list", "nomi");
      if (st.t === "N") input.setAttribute("inputmode", "decimal");
    }
    input.addEventListener("input", function () { cambia(input.value, true); });
    box.appendChild(input);
    box.appendChild(resa);
    azioni.appendChild(el("button", { classe: "link", type: "button", testo: "ometti", title: "Non scrivere nulla al posto di questa voce",
      onclick: function () { cambia(""); } }));
    if (R[chiave] !== undefined) {
      azioni.appendChild(el("button", { classe: "link", type: "button", testo: "azzera", onclick: function () { cambia(undefined); disegnaStop(); } }));
    }
    segnaScelte();
    aggiornaResa();
  }

  function rigaClausola(c, attiva, onScegli, chiaveAnteprima) {
    var riga = el("div", { classe: "clausola" + (c.manca ? " manca" : "") });
    var nome = el("button", { classe: "nome" + (attiva ? " attiva" : ""), type: "button", testo: c.nome, onclick: onScegli });
    if (c.manca) nome.disabled = true;
    riga.appendChild(nome);
    var cont = el("div");
    if (!c.manca) {
      riga.appendChild(el("button", { classe: "occhio", type: "button", title: "Mostra il testo", testo: ui.anteprime[chiaveAnteprima] ? "▾" : "▸",
        onclick: function () { ui.anteprime[chiaveAnteprima] = !ui.anteprime[chiaveAnteprima]; disegnaStop(); } }));
      if (connesso) riga.appendChild(el("button", { classe: "occhio", type: "button", title: "Modifica la clausola nel formulario", testo: "✎",
        onclick: function () { apriEditor("clausole", c.chiave); } }));
      if (ui.anteprime[chiaveAnteprima]) cont.appendChild(el("div", { classe: "anteprima-clausola", testo: testoClausola(M.doc(c.chiave)) }));
    }
    var wrap = el("div", {}, [riga, cont]);
    return wrap;
  }

  function widgetClausole(v, box, azioni) {
    var st = v.stop, chiave = v.chiaveValore, multipla = st.f === 2;
    var scelte = Array.isArray(R[chiave]) ? R[chiave] : null;
    var candidati = M.candidati(st);
    var espanso = !scelte || ui.espansi[v.percorso];

    if (scelte && scelte.length) {
      var fatte = el("div", { classe: "scelte-fatte" });
      scelte.forEach(function (k, i) {
        var d = M.doc(k);
        var riga = el("div", { classe: "scelta-fatta" }, [el("span", { testo: (multipla ? (i + 1) + ". " : "") + (d ? d.nome : k) })]);
        if (connesso && d) riga.appendChild(el("button", { type: "button", title: "Modifica la clausola nel formulario", testo: "✎", onclick: function () { apriEditor("clausole", k); } }));
        if (multipla && i > 0) {
          riga.appendChild(el("button", { type: "button", title: "Sposta su", testo: "↑", onclick: function () {
            var a = scelte.slice(); var t = a[i - 1]; a[i - 1] = a[i]; a[i] = t;
            rinumera(v.percorso, function (n) { return n === i ? i - 1 : (n === i - 1 ? i : n); });
            imposta(chiave, a); disegnaStop();
          } }));
        }
        riga.appendChild(el("button", { type: "button", title: "Togli", testo: "✕", onclick: function () {
          var a = scelte.slice(); a.splice(i, 1);
          rinumera(v.percorso, function (n) { return n === i ? null : (n > i ? n - 1 : n); });
          imposta(chiave, a.length ? a : undefined); disegnaStop();
        } }));
        fatte.appendChild(riga);
      });
      box.appendChild(fatte);
    } else if (scelte && !scelte.length) {
      box.appendChild(el("div", { classe: "resa", testo: "Nessuna clausola: la voce viene omessa." }));
    }

    if (scelte) {
      azioni.appendChild(el("button", { classe: "link", type: "button", testo: espanso && ui.espansi[v.percorso] ? "chiudi" : (multipla ? "aggiungi" : "cambia"),
        onclick: function () { ui.espansi[v.percorso] = !ui.espansi[v.percorso]; disegnaStop(); } }));
      azioni.appendChild(el("button", { classe: "link", type: "button", testo: "azzera", onclick: function () {
        rinumera(v.percorso, function () { return null; });
        imposta(chiave, undefined); ui.espansi[v.percorso] = false; disegnaStop();
      } }));
    }
    if (!espanso) return;

    var lista = el("div", { classe: "clausole" });
    candidati.forEach(function (c) {
      var attiva = !multipla && scelte && scelte[0] === c.chiave;
      lista.appendChild(rigaClausola(c, attiva, function () {
        if (multipla) {
          imposta(chiave, (scelte || []).concat([c.chiave]));
        } else {
          rinumera(v.percorso, function () { return null; });
          imposta(chiave, [c.chiave]);
          ui.espansi[v.percorso] = false;
        }
        disegnaStop();
      }, v.percorso + "|" + c.chiave));
    });
    lista.appendChild(el("div", { classe: "scelte" }, [
      el("button", { classe: "scelta" + (scelte && !scelte.length ? " attiva" : ""), type: "button", testo: "Nessuna",
        onclick: function () {
          rinumera(v.percorso, function () { return null; });
          imposta(chiave, []); ui.espansi[v.percorso] = false; disegnaStop();
        } })
    ]));
    if (multipla) box.appendChild(el("div", { classe: "resa", testo: "Puoi scegliere più clausole, anche la stessa più volte (es. due persone fisiche)." }));
    box.appendChild(lista);
  }

  // ---------------------------------------------------------------- anteprima e scrittura

  function compila() {
    return M.compila(minuta, R).map(function (p) {
      if (!opz.segnaposto) p.r = p.r.map(function (r) { if (r.hl) { r = Object.assign({}, r); r.t = ""; } return r; });
      return p;
    });
  }

  function mostraAnteprima() {
    var cont = $("testoAnteprima");
    cont.innerHTML = "";
    compila().forEach(function (p) {
      var par = el("p", { classe: p.a || "l" });
      if (p.li) par.style.marginLeft = (p.li / 20) + "pt";
      p.r.forEach(function (r) {
        if (r.br) { par.appendChild(el("br")); return; }
        var n = r.hl ? el("mark", { testo: r.t }) : document.createTextNode(r.t);
        if (r.b || r.i || r.u) {
          var s = el("span", {}, [n]);
          if (r.b) s.style.fontWeight = "bold";
          if (r.i) s.style.fontStyle = "italic";
          if (r.u) s.style.textDecoration = "underline";
          n = s;
        }
        par.appendChild(n);
      });
      if (!p.r.length) par.appendChild(el("br"));
      cont.appendChild(par);
    });
    $("velo").hidden = false;
  }

  function ooxml() {
    return Atti.ooxml(compila(), opz.formato ? { font: "Arial", dimensione: 12, interlinea: 562 } : { font: "" });
  }

  function scrivi(modo) {
    if (!inWord) { stato("Il pannello non è aperto dentro Word.", "err"); return; }
    var x;
    try { x = ooxml(); } catch (e) { stato("Errore nella composizione: " + e.message, "err"); return; }
    Word.run(function (ctx) {
      var body = ctx.document.body;
      body.load("text");
      return ctx.sync().then(function () {
        var vuoto = !String(body.text || "").trim();
        if (!vuoto && !modo) return "chiedi";
        if (vuoto || modo === "sostituisci") body.insertOoxml(x, "Replace");
        else ctx.document.getSelection().insertOoxml(x, "Replace");
        return ctx.sync().then(function () { return "fatto"; });
      });
    }).then(function (esito) {
      if (esito === "chiedi") { $("conferma").hidden = false; $("bottoniPrincipali").hidden = true; return; }
      $("conferma").hidden = true; $("bottoniPrincipali").hidden = false;
      stato("Atto scritto nel documento.", "ok");
    }).catch(function (e) {
      $("conferma").hidden = true; $("bottoniPrincipali").hidden = false;
      stato("Scrittura non riuscita" + (e && e.message ? ": " + e.message : "."), "err");
    });
  }

  // ---------------------------------------------------------------- modifica del formulario

  function chiaveNuova(sez, cat, nome) {
    var k = ((cat ? cat.replace(/\\/g, "/").replace(/^\/+|\/+$/g, "") + "/" : "") + nome + (sez === "minute" ? ".sat" : ".sap")).toLowerCase();
    return k.normalize ? k.normalize("NFC") : k;
  }

  function nomiCategorie() {
    var visti = {}, out = [];
    Object.keys(F.clausole).forEach(function (k) { var c = F.clausole[k].cat; if (c && !visti[c]) { visti[c] = 1; out.push(c); } });
    return out.sort(function (a, b) { return a.localeCompare(b); });
  }

  function elencoClausole() {
    return Object.keys(F.clausole).map(function (k) { return { chiave: k, nome: F.clausole[k].nome, cat: F.clausole[k].cat }; });
  }

  function dopoModifica(v) {
    Object.assign(infoMod, Archivio.applica(F, [v]));
    M = Atti.crea(F);
    aggiornaVersione(true);
  }

  function apriEditor(sez, chiave) {
    if (!connesso) { stato("Per modificare il formulario serve il collegamento a SharePoint.", "err"); return; }
    var doc = chiave ? F[sez][chiave] : null;
    var info = chiave ? infoMod[sez + "|" + chiave] : null;
    Editor.apri($("editor"), {
      sez: sez, chiave: chiave, doc: doc, nuovo: !chiave, info: info,
      categorie: sez === "clausole" ? nomiCategorie() : Object.keys(F.minute).map(function (k) { return F.minute[k].cat; }).filter(function (c, i, a) { return c && a.indexOf(c) === i; }),
      clausole: elencoClausole(),
      salva: function (nuovoDoc) {
        var k = chiave;
        if (!k) {
          k = chiaveNuova(sez, nuovoDoc.cat, nuovoDoc.nome);
          if (F[sez][k]) return Promise.reject(new Error("Esiste già una " + (sez === "minute" ? "minuta" : "clausola") + " con questo nome in questa categoria."));
        }
        var eTag = infoMod[sez + "|" + k] ? infoMod[sez + "|" + k].eTag : null;
        return Archivio.salva(sez, k, nuovoDoc, eTag).then(function (v) {
          chiave = k;
          dopoModifica(v);
          info = infoMod[sez + "|" + k];
          disegnaTutto();
        });
      },
      elimina: function () {
        var eTag = infoMod[sez + "|" + chiave] ? infoMod[sez + "|" + chiave].eTag : null;
        return Archivio.salva(sez, chiave, null, eTag).then(function (v) {
          dopoModifica(v);
          if (sez === "minute" && minuta === chiave) minuta = null;
          disegnaTutto();
        });
      },
      ripristina: function () {
        return Archivio.ripristina(info.id).then(function () { return daSharePoint(false); }).then(function () { disegnaTutto(); });
      },
      chiudi: function () { disegnaTutto(); }
    });
  }

  function disegnaArchivio() {
    if (!F || !$("secArchivio").open) return;
    var cont = $("risultatiArchivio");
    cont.innerHTML = "";
    var parole = $("cercaArchivio").value.toLowerCase().split(/\s+/).filter(Boolean);
    var modificate = Object.keys(infoMod).length;
    $("notaArchivio").textContent = modificate ? modificate + (modificate === 1 ? " voce modificata" : " voci modificate") + " rispetto al formulario convertito da RedAtto." : "";
    if (!parole.length) {
      // senza ricerca: le ultime modificate
      Object.keys(infoMod).slice(-30).reverse().forEach(function (k) {
        var p = k.split("|"), d = F[p[0]][p[1]], im = infoMod[k];
        cont.appendChild(rigaArchivio(p[0], p[1], d ? d.nome : p[1].split("/").pop(), (d ? d.cat : "") + (im.eliminata ? " · eliminata" : " · modificata" + (im.autore ? " da " + im.autore : "")), !d));
      });
      if (!modificate) cont.appendChild(el("div", { classe: "vuota", testo: "Scrivi per cercare tra " + Object.keys(F.clausole).length + " clausole e " + Object.keys(F.minute).length + " minute." }));
      return;
    }
    var n = 0;
    ["minute", "clausole"].forEach(function (sez) {
      Object.keys(F[sez]).forEach(function (k) {
        if (n >= 60) return;
        var d = F[sez][k], t = (d.nome + " " + d.cat).toLowerCase();
        if (!parole.every(function (p) { return t.indexOf(p) >= 0; })) return;
        n++;
        cont.appendChild(rigaArchivio(sez, k, d.nome, (sez === "minute" ? "Minuta · " : "") + d.cat + (infoMod[sez + "|" + k] ? " · modificata" : ""), false));
      });
    });
    if (!n) cont.appendChild(el("div", { classe: "vuota", testo: "Nessun risultato." }));
  }

  function rigaArchivio(sez, k, nome, sotto, eliminata) {
    return el("button", { classe: "minuta", type: "button", onclick: function () {
      if (eliminata) { stato("Questa voce è stata eliminata: per riaverla usa \"Torna al testo originale\" dopo averla ricreata, oppure ripristinala da SharePoint.", "err"); return; }
      apriEditor(sez, k);
    } }, [el("div", { testo: "✎ " + nome }), el("div", { classe: "file-nome", testo: sotto })]);
  }

  // ---------------------------------------------------------------- disegno generale

  function disegnaTutto() {
    $("secFormulario").hidden = !!F || !$("secAccesso").hidden;
    if (F) $("secAccesso").hidden = true;
    $("secArchivio").hidden = !F || !connesso;
    $("modificaMinuta").hidden = !connesso;
    $("secMinuta").hidden = !F;
    var conMinuta = !!(F && minuta && F.minute[minuta]);
    $("minutaScelta").hidden = !conMinuta;
    $("minutaElenco").hidden = conMinuta;
    $("secStop").hidden = !conMinuta;
    $("barra").hidden = !conMinuta;
    if (!F) return;
    if (conMinuta) {
      $("minutaNome").textContent = F.minute[minuta].nome;
      $("minutaCat").textContent = F.minute[minuta].cat;
      disegnaStop();
    } else disegnaMinute();
  }

  function collega() {
    var zona = $("zona"), input = $("file");
    input.addEventListener("change", function () { caricaFile(input.files[0]); input.value = ""; });
    zona.addEventListener("keydown", function (e) { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); input.click(); } });
    ["dragenter", "dragover"].forEach(function (ev) { document.addEventListener(ev, function (e) { e.preventDefault(); zona.classList.add("sopra"); }); });
    ["dragleave", "drop"].forEach(function (ev) { document.addEventListener(ev, function (e) { e.preventDefault(); zona.classList.remove("sopra"); }); });
    document.addEventListener("drop", function (e) {
      var f = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
      if (f && /\.json$/i.test(f.name)) caricaFile(f);
    });
    $("aggiornaFormulario").addEventListener("click", function () { input.click(); });
    $("cercaMinuta").addEventListener("input", disegnaMinute);
    $("cambiaMinuta").addEventListener("click", function () { minuta = null; disegnaTutto(); $("cercaMinuta").focus(); });
    $("cercaStop").addEventListener("input", function () { ui.filtro = $("cercaStop").value; disegnaStop(); });
    $("soloDaFare").addEventListener("change", function () { ui.soloDaFare = $("soloDaFare").checked; disegnaStop(); });
    $("optFormato").checked = opz.formato;
    $("optSegnaposto").checked = opz.segnaposto;
    var salvaOpz = function () {
      opz.formato = $("optFormato").checked; opz.segnaposto = $("optSegnaposto").checked;
      try { window.localStorage.setItem(CHIAVE_OPZ, JSON.stringify(opz)); } catch (e) { /* ignora */ }
    };
    $("optFormato").addEventListener("change", salvaOpz);
    $("optSegnaposto").addEventListener("change", salvaOpz);
    $("anteprima").addEventListener("click", mostraAnteprima);
    $("chiudiVelo").addEventListener("click", function () { $("velo").hidden = true; });
    $("scrivi").addEventListener("click", function () { scrivi(null); });
    $("sostituisci").addEventListener("click", function () { scrivi("sostituisci"); });
    $("alCursore").addEventListener("click", function () { scrivi("cursore"); });
    $("annulla").addEventListener("click", function () { $("conferma").hidden = true; $("bottoniPrincipali").hidden = false; });

    $("accedi").addEventListener("click", function () {
      stato("Accesso in corso…");
      daSharePoint(true).then(function () {
        $("secAccesso").hidden = true; stato("Collegato al formulario dello studio.", "ok");
        ripristinaDaDocumento(); disegnaTutto();
      }).catch(function (e) { stato("Accesso non riuscito: " + (e && e.message ? e.message : e), "err"); });
    });
    $("modificaMinuta").addEventListener("click", function () { apriEditor("minute", minuta); });
    $("cercaArchivio").addEventListener("input", disegnaArchivio);
    $("secArchivio").addEventListener("toggle", disegnaArchivio);
    $("nuovaClausola").addEventListener("click", function () { apriEditor("clausole", null); });
    $("nuovaMinuta").addEventListener("click", function () { apriEditor("minute", null); });

    avviaFormulario().then(function () {
      ripristinaDaDocumento();
      disegnaTutto();
    });
  }

  var pronto = { dom: false, office: !(window.Office && Office.onReady) };
  function avvia() { if (pronto.dom && pronto.office) collega(); }
  document.addEventListener("DOMContentLoaded", function () { pronto.dom = true; avvia(); });
  if (window.Office && Office.onReady) {
    Office.onReady(function (info) {
      inWord = !!(info && info.host === Office.HostType.Word);
      pronto.office = true; avvia();
    });
  }
})();
