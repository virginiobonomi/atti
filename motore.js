/*
 * Atti — motore: dalle minute e clausole del formulario (convertite da RedAtto) all'atto compilato.
 * Funziona sia nel pannello di Word sia in Node (per le prove).
 *
 * Formulario: { clausole: {chiave: doc}, categorie: {cat: [chiavi]}, minute: {chiave: doc} }
 * doc: { nome, cat, p: [paragrafi], s: {id: stop} }
 * paragrafo: { a: "j|c|r" (assente = sinistra), li, fi, x: [ {t, b, i, u, k} | {s: id, l: etichetta} | {br: 1} ] }
 * stop: { n: nome, t: tipo (T, LC, D, N, CM, CF, VE, ...), f: formato, k: sottotipo, ft, v: valori, pr, po, va, sv, co }
 *
 * Risposte: { percorso: valore }. Percorso di uno stop: prefisso + id. Le clausole scelte in una lista
 * hanno prefisso  <percorso stop>[<n>]/ . Valore di una lista clausole: array di chiavi di clausola.
 * Le date "DATA STIPULA" (e gli stop speciali #...) sono condivise in tutto l'atto con chiave "@" + nome.
 */
(function (root) {
  "use strict";

  var NL = (typeof require !== "undefined" && typeof module !== "undefined") ? require("./numeri.js") : root.NumeriInLettere;

  var MESI = ["gennaio", "febbraio", "marzo", "aprile", "maggio", "giugno", "luglio",
    "agosto", "settembre", "ottobre", "novembre", "dicembre"];

  // ---------------------------------------------------------------- formattazione dei valori

  function lettere(n) { return NL.interoInLettere(String(n)); }

  function caso(s, ft) {
    if (!s || !ft) return s;
    if (ft === "Minuscolo") return s.toLowerCase();
    if (ft === "MAIUSCOLO") return s.toUpperCase();
    if (ft === "InizialiMaiuscole") {
      return s.toLowerCase().replace(/(^|[\s'(\-\/])([a-zà-ÿ])/g, function (m, a, b) { return a + b.toUpperCase(); });
    }
    return s;
  }

  // "2026-11-08" (campo data) oppure "8/11/2026"
  function leggiData(v) {
    var m = String(v || "").trim().match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
    if (m) return { g: +m[3], m: +m[2], a: +m[1] };
    m = String(v || "").trim().match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{2,4})$/);
    if (m) return { g: +m[1], m: +m[2], a: m[3].length === 2 ? 2000 + +m[3] : +m[3] };
    return null;
  }

  // formati data di RedAtto (manuale, par. 11.5)
  function formattaData(v, f) {
    var d = leggiData(v);
    if (!d || d.m < 1 || d.m > 12) return String(v || "");
    var g = String(d.g), mm = ("0" + d.m).slice(-2), aaaa = String(d.a), aa = aaaa.slice(-2);
    var giorno = d.g === 1 ? "primo" : lettere(d.g), mese = MESI[d.m - 1], anno = lettere(d.a);
    switch (f) {
      case 1: return g + "/" + mm + "/" + aaaa;
      case 2: return giorno + " " + mese + " " + anno;
      case 3: return g + " (" + giorno + ") " + mm + " (" + mese + ") " + aaaa + " (" + anno + ")";
      case 4: return g + "/" + mm + "/" + aa;
      case 5: return g + " " + mese + " " + aaaa;
      case 6: return g + " (" + giorno + ") " + mese + " " + aaaa + " (" + anno + ")";
      case 7: return g + " " + mese + " " + aa;
      case 8: return g;
      case 9: return giorno;
      case 10: return String(d.m);
      case 11: return mese;
      case 12: return aa;
      case 13: return aaaa;
      case 14: return anno;
      case 15: return giorno + " " + mese + " " + anno + " (" + g + "." + mm + "." + aaaa + ")";
      default: return g + " " + mese + " " + aaaa;
    }
  }

  function leggiNumero(v) {
    var t = String(v || "").trim().replace(/^(euro|€)\s*/i, "").replace(/\s/g, "");
    if (!t) return null;
    var intero, dec = "";
    if (t.indexOf(",") >= 0) { var p = t.split(","); if (p.length !== 2) return null; intero = p[0].replace(/\./g, ""); dec = p[1]; }
    else if (/^\d{1,3}(\.\d{3})+$/.test(t)) intero = t.replace(/\./g, "");
    else if (/^\d+\.\d+$/.test(t)) { var q = t.split("."); intero = q[0]; dec = q[1]; }
    else intero = t;
    if (!/^\d+$/.test(intero || "0") || (dec && !/^\d+$/.test(dec))) return null;
    return { intero: (intero || "0").replace(/^0+(?=\d)/, ""), dec: dec };
  }

  function puntini(s) { return s.replace(/\B(?=(\d{3})+(?!\d))/g, "."); }

  function decInLettere(dec) {
    var m = dec.match(/^(0*)(\d*)$/), parole = [];
    for (var i = 0; i < m[1].length; i++) parole.push("zero");
    if (m[2]) parole.push(lettere(m[2]));
    return parole.join(" ");
  }

  // formati numero di RedAtto (manuale, par. 11.5); k = "CU" per gli importi
  function formattaNumero(v, f, k) {
    var n = leggiNumero(v);
    if (!n) return String(v || "");
    var dec = n.dec;
    if (k === "CU" && dec.length === 1) dec += "0";
    var cifre = n.intero + (dec ? "," + dec : "");
    var cifreP = puntini(n.intero) + (dec ? "," + dec : "");
    var lett = lettere(n.intero) + (dec ? " virgola " + decInLettere(dec) : "");
    switch (f) {
      case 1: return cifre;
      case 2: return cifreP;
      case 3: return cifre + " (" + lett + ")";
      case 4: return cifreP + " (" + lett + ")";
      case 5: return "(" + lett + ") " + cifre;
      case 6: return "(" + lett + ") " + cifreP;
      case 10: return lettere(n.intero);
      case 11: return dec ? decInLettere(dec) : "zero";
      case 12: return cifreP + " (" + lettere(n.intero) + " euro" + (dec && +dec ? " e " + lettere(+dec) + " centesimi" : "") + ")";
      case 13: return lett + " (" + cifreP + ")";
      case 7: case 8: case 9: return lett;
      default: return f >= 14 ? lett : cifreP;
    }
  }

  // valore scritto nell'atto per uno stop compilato (vuoto = omesso)
  function formatta(st, v) {
    if (v === undefined || v === null) return null;
    v = String(v);
    if (v.trim() === "") return "";
    var out;
    if (st.t === "D") out = formattaData(v, st.f);
    else if (st.t === "N") out = formattaNumero(v, st.f, st.k);
    else if (st.t === "CF") out = v.toUpperCase().replace(/\s+/g, "");
    else out = caso(v, st.ft);
    return (st.pr || "") + out + (st.po || "");
  }

  // ---------------------------------------------------------------- formulario

  function crea(F) {
    var nomiCat = {};

    function doc(chiave) { return F.clausole[chiave] || null; }

    // clausole offerte da uno stop "lista clausole"
    function candidati(st) {
      var out = [], visti = {};
      (st.v || []).forEach(function (coppia) {
        var cat = coppia[0].replace(/^\/+|\/+$/g, ""), f = coppia[1];
        var lista = f === "*" ? (F.categorie[cat] || []) : [f];
        if (f === "*" && !F.categorie[cat]) out.push({ chiave: null, nome: "Categoria «" + cat + "» non presente nel formulario", manca: true });
        lista.forEach(function (k) {
          if (visti[k]) return;
          visti[k] = 1;
          var d = F.clausole[k];
          out.push(d ? { chiave: k, nome: d.nome } : { chiave: k, nome: k.split("/").pop().replace(/\.sap$/, "") + " (non presente nel formulario)", manca: true });
        });
      });
      return out;
    }

    function chiaveCondivisa(st) {
      if (st.t === "D" && /^DATA STIPULA$/i.test(st.n.trim())) return "@DATA STIPULA";
      if (/^#/.test(st.n.trim())) return "@" + st.n.trim();
      return null;
    }

    function percorsoValore(d, prefisso, id) {
      var st = d.s[id];
      var cond = st && chiaveCondivisa(st);
      if (cond) return cond;
      // variabile copia: stesso valore dello stop originale
      if (st && st.co && d.s[st.co]) return prefisso + st.co;
      return prefisso + id;
    }

    // valore predefinito (stop con "stampa valore")
    function predefinito(st) {
      if (st.sv && st.va && st.va.indexOf("/") < 0) return st.va;
      return undefined;
    }

    function valore(R, d, prefisso, id) {
      var p = percorsoValore(d, prefisso, id);
      if (Object.prototype.hasOwnProperty.call(R, p)) return R[p];
      return predefinito(d.s[id] || {});
    }

    // ------------------------------------------------ elenco degli stop (per il pannello)

    function testoParagrafo(par, d, prefisso, R, finoA) {
      var s = "";
      for (var i = 0; i < par.x.length && (finoA === undefined || i < finoA); i++) {
        var x = par.x[i];
        if (x.t !== undefined) s += x.t;
        else if (x.s !== undefined) {
          var st = d.s[x.s] || {};
          var val = st.t === "LC" ? null : formatta(st, valore(R, d, prefisso, x.s));
          s += val !== null && val !== undefined ? val : "[" + x.l.trim() + "]";
        } else s += " ";
      }
      return s;
    }

    function elenco(chiaveMinuta, R) {
      var m = F.minute[chiaveMinuta];
      var out = [];
      if (!m) return out;
      function giro(d, prefisso, livello, nomeClausola) {
        var sezione = "";
        d.p.forEach(function (par) {
          var haStop = par.x.some(function (x) { return x.s !== undefined; });
          if (!haStop) {
            var t = par.x.map(function (x) { return x.t || ""; }).join("").trim();
            if (t && t.length <= 70 && livello === 0 && !/^[°\s]+$/.test(t)) sezione = t;
            return;
          }
          par.x.forEach(function (x, i) {
            if (x.s === undefined) return;
            var st = d.s[x.s] || { n: x.l, t: "T", f: 1 };
            var p = prefisso + x.s;
            var prima = testoParagrafo(par, d, prefisso, R, i);
            var dopo = "";
            for (var j = i + 1; j < par.x.length && dopo.length < 40; j++) dopo += par.x[j].t !== undefined ? par.x[j].t : (par.x[j].s !== undefined ? "[" + par.x[j].l.trim() + "]" : " ");
            var pv = percorsoValore(d, prefisso, x.s);
            var voce = {
              percorso: p, chiaveValore: pv, stop: st, etichetta: x.l.trim() || st.n, livello: livello,
              sezione: sezione, clausola: nomeClausola || "", prima: prima.slice(-70), dopo: dopo.slice(0, 40),
              copia: pv !== p && pv.charAt(0) !== "@" , condiviso: pv.charAt(0) === "@"
            };
            var v = valore(R, d, prefisso, x.s);
            voce.valore = v;
            voce.compilato = st.t === "LC" ? Array.isArray(v) : (v !== undefined && v !== null);
            out.push(voce);
            if (st.t === "LC" && Array.isArray(v)) {
              v.forEach(function (k, n) {
                var c = F.clausole[k];
                if (c) giro(c, p + "[" + n + "]/", livello + 1, c.nome);
              });
            }
          });
        });
      }
      giro(m, "", 0, "");
      // per le variabili copia e i valori condivisi si mostra solo la prima occorrenza
      var visti = {};
      out.forEach(function (v) {
        if (visti[v.chiaveValore] && v.stop.t !== "LC") v.duplicato = true;
        visti[v.chiaveValore] = 1;
      });
      return out;
    }

    // ------------------------------------------------ atto compilato

    // paragrafi in uscita: { a, li, fi, r: [ {t, b, i, u, k, hl} | {br:1} ] }
    function compila(chiaveMinuta, R, opz) {
      opz = opz || {};
      var m = F.minute[chiaveMinuta];
      if (!m) return [];
      var profondita = 0;

      function espandi(d, prefisso) {
        if (++profondita > 40) throw new Error("Clausole annidate troppo in profondità");
        var paragrafi = [];
        d.p.forEach(function (par) {
          var corrente = { a: par.a || "l", li: par.li || 0, fi: par.fi || 0, r: [], conStop: false };
          var lista = [corrente];
          par.x.forEach(function (x) {
            var att = lista[lista.length - 1];
            if (x.t !== undefined) { att.r.push(copiaRun(x)); return; }
            if (x.br) { att.r.push({ br: 1 }); return; }
            att.conStop = true;
            var st = d.s[x.s] || { n: x.l, t: "T" };
            var v = valore(R, d, prefisso, x.s);
            if (st.t === "LC") {
              if (!Array.isArray(v)) { att.r.push(segnaposto(x)); return; }
              var p = prefisso + x.s;
              var inseriti = [];
              v.forEach(function (k, n) {
                var c = F.clausole[k];
                if (!c) return;
                inseriti = inseriti.concat(espandi(c, p + "[" + n + "]/"));
              });
              if (!inseriti.length) return;
              // la prima riga della clausola prosegue il paragrafo, l'ultima accoglie il resto
              att.r = att.r.concat(inseriti[0].r);
              for (var q = 1; q < inseriti.length; q++) {
                var np = inseriti[q];
                np.conStop = true;
                lista.push(np);
              }
              return;
            }
            var testo = formatta(st, v);
            if (testo === null || testo === undefined) { att.r.push(segnaposto(x)); return; }
            if (testo === "") return;
            var run = { t: testo };
            if (x.b) run.b = 1;
            if (x.i) run.i = 1;
            if (x.u) run.u = 1;
            att.r.push(run);
          });
          lista.forEach(function (pp) { paragrafi.push(pp); });
        });
        profondita--;
        return paragrafi;
      }

      function copiaRun(x) {
        var r = { t: x.t };
        ["b", "i", "u", "k"].forEach(function (k) { if (x[k]) r[k] = 1; });
        return r;
      }

      function segnaposto(x) {
        var m = x.l.match(/^(\s*)([\s\S]*?)(\s*)$/);
        return { t: m[1] + "[" + m[2] + "]" + m[3], hl: 1 };
      }

      var out = espandi(m, "");
      // un paragrafo rimasto vuoto perché lo stop è stato omesso si toglie
      out = out.filter(function (p) {
        if (!p.conStop) return true;
        return p.r.some(function (r) { return r.br || (r.t && r.t.trim()); });
      });
      // pulizia: spazi doppi creati dagli stop omessi
      out.forEach(function (p) {
        for (var i = 0; i < p.r.length - 1; i++) {
          var a = p.r[i], b = p.r[i + 1];
          if (a.t && b.t && / $/.test(a.t) && /^ /.test(b.t)) b.t = b.t.replace(/^ +/, "");
        }
        delete p.conStop;
      });
      return out;
    }

    function testo(paragrafi) {
      return paragrafi.map(function (p) {
        return p.r.map(function (r) { return r.br ? "\n" : r.t; }).join("");
      }).join("\n");
    }

    return {
      F: F, candidati: candidati, elenco: elenco, compila: compila, testo: testo, doc: doc,
      minute: function () {
        return Object.keys(F.minute).map(function (k) { return { chiave: k, nome: F.minute[k].nome, cat: F.minute[k].cat }; });
      }
    };
  }

  // ---------------------------------------------------------------- OOXML per Word

  function esc(s) {
    return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  }

  function ooxml(paragrafi, opz) {
    opz = opz || {};
    var font = opz.font === undefined ? "Arial" : opz.font;
    var dim = opz.dimensione || 12;
    var rprBase = (font ? '<w:rFonts w:ascii="' + esc(font) + '" w:hAnsi="' + esc(font) + '" w:cs="' + esc(font) + '"/>' : "") +
      (font ? '<w:sz w:val="' + dim * 2 + '"/><w:szCs w:val="' + dim * 2 + '"/>' : "");
    var corpo = paragrafi.map(function (p) {
      var ppr = "";
      if (opz.interlinea) ppr += '<w:spacing w:before="0" w:after="0" w:line="' + opz.interlinea + '" w:lineRule="exact"/>';
      else ppr += '<w:spacing w:before="0" w:after="0"/>';
      if (p.li || p.fi) ppr += '<w:ind w:left="' + (p.li || 0) + '"' + (p.fi < 0 ? ' w:hanging="' + (-p.fi) + '"' : (p.fi ? ' w:firstLine="' + p.fi + '"' : "")) + "/>";
      var jc = { j: "both", c: "center", r: "right" }[p.a];
      if (jc) ppr += '<w:jc w:val="' + jc + '"/>';
      var runs = p.r.map(function (r) {
        if (r.br) return "<w:r><w:br/></w:r>";
        if (!r.t) return "";
        var rpr = rprBase + (r.b ? "<w:b/>" : "") + (r.i ? "<w:i/>" : "") + (r.k ? "<w:smallCaps/>" : "") +
          (r.u ? '<w:u w:val="single"/>' : "") + (r.hl ? '<w:highlight w:val="yellow"/>' : "");
        return r.t.split("\t").map(function (pezzo, i) {
          return (i ? "<w:r>" + (rpr ? "<w:rPr>" + rpr + "</w:rPr>" : "") + "<w:tab/></w:r>" : "") +
            (pezzo ? "<w:r>" + (rpr ? "<w:rPr>" + rpr + "</w:rPr>" : "") + '<w:t xml:space="preserve">' + esc(pezzo) + "</w:t></w:r>" : "");
        }).join("");
      }).join("");
      return "<w:p><w:pPr>" + ppr + (rprBase ? "<w:rPr>" + rprBase + "</w:rPr>" : "") + "</w:pPr>" + runs + "</w:p>";
    }).join("");
    return '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">' +
      '<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
      "</Relationships></pkg:xmlData></pkg:part>" +
      '<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData>' +
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' + corpo + "</w:body></w:document>" +
      "</pkg:xmlData></pkg:part></pkg:package>";
  }

  var api = { crea: crea, ooxml: ooxml, formatta: formatta, formattaData: formattaData, formattaNumero: formattaNumero };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.Atti = api;
})(this);
