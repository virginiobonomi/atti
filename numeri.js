/*
 * Numeri in lettere — motore di conversione (italiano, formule notarili)
 * Converte numeri, importi, percentuali, date e quote/frazioni in lettere.
 * Funziona sia dentro Word (componente aggiuntivo) sia in un normale browser.
 */
(function (root) {
  "use strict";

  var UNITA = ["zero", "uno", "due", "tre", "quattro", "cinque", "sei", "sette", "otto", "nove",
    "dieci", "undici", "dodici", "tredici", "quattordici", "quindici", "sedici",
    "diciassette", "diciotto", "diciannove"];
  var DECINE = ["", "", "venti", "trenta", "quaranta", "cinquanta", "sessanta", "settanta", "ottanta", "novanta"];
  var MESI = ["gennaio", "febbraio", "marzo", "aprile", "maggio", "giugno", "luglio",
    "agosto", "settembre", "ottobre", "novembre", "dicembre"];

  // 1..99
  function sotto100(n) {
    if (n < 20) return UNITA[n];
    var d = Math.floor(n / 10), u = n % 10;
    var dec = DECINE[d];
    if (u === 0) return dec;
    if (u === 1 || u === 8) dec = dec.slice(0, -1); // ventuno, ventotto
    return dec + UNITA[u];
  }

  // 1..999
  function sotto1000(n) {
    var c = Math.floor(n / 100), r = n % 100, s = "";
    if (c > 0) s = (c === 1) ? "cento" : UNITA[c] + "cento";
    if (r > 0) {
      var t = sotto100(r);
      if (c > 0 && t.indexOf("ott") === 0) s = s.slice(0, -1); // centotto, centottanta
      s += t;
    }
    return s;
  }

  // ventitre -> ventitré, centotre -> centotré, milletre -> milletré (non "tre" da solo)
  function accenta(s) {
    return s.split(" ").map(function (w) {
      return /[a-z]tre$/.test(w) ? w.slice(0, -3) + "tré" : w;
    }).join(" ");
  }

  // Parte intera (stringa di cifre) in lettere
  function interoInLettere(cifre, opz) {
    opz = opz || {};
    var str = String(cifre).replace(/^0+(?=\d)/, "");
    if (!/^\d+$/.test(str)) throw new Error("Numero non valido");
    if (str === "0") return "zero";
    if (str.length > 12) throw new Error("Numero troppo grande (massimo 999 miliardi)");
    var n = ("000000000000" + str).slice(-12);
    var mld = +n.slice(0, 3), mln = +n.slice(3, 6), mig = +n.slice(6, 9), un = +n.slice(9);
    var parti = [];
    if (mld) parti.push(mld === 1 ? "un miliardo" : sotto1000(mld) + " miliardi");
    if (mln) parti.push(mln === 1 ? "un milione" : sotto1000(mln) + " milioni");
    var resto = "";
    if (mig) resto += (mig === 1) ? "mille" : sotto1000(mig) + "mila";
    if (un) resto += sotto1000(un);
    if (resto) parti.push(resto);
    var testo = accenta(parti.join(" "));
    return opz.milioniUniti ? testo.replace(/ /g, "") : testo;
  }

  // Decimali: gli zeri iniziali si leggono "zero" (",05" = "zero cinque")
  function decimaliInLettere(dec, opz) {
    var m = dec.match(/^(0*)(\d*)$/);
    var parole = [];
    for (var i = 0; i < m[1].length; i++) parole.push("zero");
    if (m[2]) parole.push(interoInLettere(m[2], opz));
    return parole.join(" ");
  }

  function formattaCifre(intero) {
    return intero.replace(/^0+(?=\d)/, "").replace(/\B(?=(\d{3})+(?!\d))/g, ".");
  }

  // Riconosce "125.000,50", "125000,5", "1.500", "125000.50", "€ 1.000", "euro 300"
  function analizzaNumero(testo) {
    var t = String(testo).trim()
      .replace(/^(euro|eur|€)\s*/i, "").replace(/\s*(euro|eur|€)$/i, "")
      .replace(/[\s  ']/g, "");
    var perc = false;
    if (/%$/.test(t)) { perc = true; t = t.slice(0, -1); }
    var neg = false;
    if (/^[-−]/.test(t)) { neg = true; t = t.slice(1); }
    if (t.charAt(0) === "+") t = t.slice(1);
    if (!/^[\d.,]+$/.test(t)) return null;
    var intero, dec = "";
    if (t.indexOf(",") >= 0) {
      var p = t.split(",");
      if (p.length !== 2) return null;
      intero = p[0]; dec = p[1];
      if (intero.indexOf(".") >= 0) {
        if (!/^\d{1,3}(\.\d{3})+$/.test(intero)) return null;
        intero = intero.replace(/\./g, "");
      }
    } else if (t.indexOf(".") >= 0) {
      if (/^\d{1,3}(\.\d{3})+$/.test(t)) intero = t.replace(/\./g, "");
      else {
        var q = t.split(".");
        if (q.length !== 2) return null;
        intero = q[0]; dec = q[1];
      }
    } else intero = t;
    if (intero === "") intero = "0";
    if (!/^\d+$/.test(intero) || (dec && !/^\d+$/.test(dec))) return null;
    return { neg: neg, intero: intero, dec: dec, perc: perc };
  }

  function numeroInLettere(num, opz, decForzati) {
    var dec = decForzati !== undefined ? decForzati : num.dec;
    var s = interoInLettere(num.intero, opz);
    if (dec) s += " virgola " + decimaliInLettere(dec, opz);
    if (num.neg) s = "meno " + s;
    if (num.perc) s += " per cento";
    return s;
  }

  function convertiNumero(testo, opz) {
    var num = analizzaNumero(testo);
    if (!num) throw new Error("Non riconosco un numero in «" + testo + "»");
    var formato = opz.formatoNumero || "cifre";
    var segno = num.neg ? "-" : "";
    if (formato === "euro" || formato === "euro-barra") {
      if (num.perc) throw new Error("Una percentuale non può essere scritta come importo in euro");
      var dec = num.dec;
      if (dec.length === 1) dec += "0";                 // 125,5 -> 125,50
      if (!dec && opz.decimaliEuro === true) dec = "00";
      var cifre = segno + formattaCifre(num.intero) + (dec ? "," + dec : "");
      var lettere;
      if (formato === "euro-barra") {
        lettere = (num.neg ? "meno " : "") + interoInLettere(num.intero, opz) + (dec ? "/" + dec : "");
      } else {
        lettere = numeroInLettere(num, opz, dec);
      }
      return { tipo: "importo", completo: "euro " + cifre + " (" + lettere + ")", lettere: lettere };
    }
    var lett = numeroInLettere(num, opz);
    var cif = segno + formattaCifre(num.intero) + (num.dec ? "," + num.dec : "") + (num.perc ? "%" : "");
    return {
      tipo: num.perc ? "percentuale" : "numero",
      completo: formato === "lettere" ? lett : cif + " (" + lett + ")",
      lettere: lett
    };
  }

  // Date: 07/11/2026, 7.11.2026, 7-11-2026, 7 novembre 2026
  function analizzaData(testo) {
    var t = String(testo).trim().toLowerCase().replace(/°/g, "");
    var g, m, a, r;
    if ((r = t.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{4})$/))) {
      g = +r[1]; m = +r[2]; a = +r[3];
    } else if ((r = t.match(/^(\d{1,2})\s+([a-z]+)\s+(\d{4})$/))) {
      g = +r[1]; m = MESI.indexOf(r[2]) + 1; a = +r[3];
      if (m === 0) return null;
    } else return null;
    var d = new Date(a, m - 1, g);
    if (d.getFullYear() !== a || d.getMonth() !== m - 1 || d.getDate() !== g) {
      throw new Error("La data «" + testo + "» non esiste");
    }
    return { g: g, m: m, a: a };
  }

  function convertiData(testo, opz) {
    var d = analizzaData(testo);
    if (!d) throw new Error("Non riconosco una data in «" + testo + "»");
    var giorno = d.g === 1 ? "primo" : interoInLettere(String(d.g));
    var mese = MESI[d.m - 1];
    var anno = interoInLettere(String(d.a));
    var lettere = giorno + " " + mese + " " + anno;
    var formato = opz.formatoData || "lettere";
    var completo;
    if (formato === "notarile") completo = anno + ", il giorno " + giorno + " del mese di " + mese;
    else if (formato === "cifre") {
      completo = ("0" + d.g).slice(-2) + "/" + ("0" + d.m).slice(-2) + "/" + d.a + " (" + lettere + ")";
    } else completo = lettere;
    return { tipo: "data", completo: completo, lettere: lettere };
  }

  // Quote e frazioni: 1/2 un mezzo, 3/4 tre quarti, 125/1000 centoventicinque millesimi
  var DENOM = {
    1: ["intero", "interi"], 2: ["mezzo", "mezzi"], 3: ["terzo", "terzi"], 4: ["quarto", "quarti"],
    5: ["quinto", "quinti"], 6: ["sesto", "sesti"], 7: ["settimo", "settimi"], 8: ["ottavo", "ottavi"],
    9: ["nono", "noni"], 10: ["decimo", "decimi"]
  };

  function denominatore(den, plurale) {
    if (DENOM[den]) return DENOM[den][plurale ? 1 : 0];
    var c = interoInLettere(String(den)).replace(/^un /, "").replace(/ /g, "").replace(/é/g, "e");
    var ord;
    if (/mila$/.test(c)) ord = c.slice(0, -4) + "millesimo";      // duemila -> duemillesimo
    else if (/(tre|sei)$/.test(c)) ord = c + "esimo";            // ventitreesimo, ventiseiesimo
    else ord = c.slice(0, -1) + "esimo";                          // undicesimo, centesimo, millesimo
    return plurale ? ord.slice(0, -1) + "i" : ord;
  }

  function convertiQuota(testo, opz) {
    var r = String(testo).trim().match(/^(\d+)\s*\/\s*(\d+)$/);
    if (!r) throw new Error("Non riconosco una quota in «" + testo + "»");
    var num = +r[1], den = +r[2];
    if (den === 0) throw new Error("Il denominatore non può essere zero");
    var lettere = num === 1 ? "un " + denominatore(den, false)
      : interoInLettere(r[1]) + " " + denominatore(den, true);
    var completo = opz.formatoNumero === "lettere" ? lettere : num + "/" + den + " (" + lettere + ")";
    return { tipo: "quota", completo: completo, lettere: lettere };
  }

  function riconosciTipo(testo) {
    var t = String(testo).trim();
    if (/^\d{1,2}[\/.\-]\d{1,2}[\/.\-]\d{4}$/.test(t) || /^\d{1,2}°?\s+[a-zA-Z]+\s+\d{4}$/.test(t)) return "data";
    if (/^\d+\s*\/\s*\d+$/.test(t)) return "quota";
    return "numero";
  }

  /**
   * converti(testo, opzioni) -> { tipo, completo, lettere }
   * opzioni: tipo ("auto"|"numero"|"data"|"quota"),
   *          formatoNumero ("lettere"|"cifre"|"euro"|"euro-barra"),
   *          formatoData ("lettere"|"notarile"|"cifre"),
   *          decimaliEuro (bool, default false), milioniUniti (bool)
   */
  function converti(testo, opzioni) {
    var opz = opzioni || {};
    var t = String(testo || "").replace(/[\r\n\t]+/g, " ").trim();
    if (!t) throw new Error("Seleziona o scrivi un numero, una data o una quota");
    var tipo = (!opz.tipo || opz.tipo === "auto") ? riconosciTipo(t) : opz.tipo;
    if (tipo === "data") return convertiData(t, opz);
    if (tipo === "quota") return convertiQuota(t, opz);
    return convertiNumero(t, opz);
  }

  /**
   * estraiFinale(testoPrimaDelCursore) -> { testo, spazi } oppure null
   * Individua il numero, la data o la quota che precede il cursore.
   * "spazi" sono gli eventuali spazi tra il numero e il cursore.
   */
  function estraiFinale(testoPrima) {
    var t = String(testoPrima || "").replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, " ");
    var m = t.match(/(\d{1,2}°?\s+([A-Za-zÀ-ÿ]+)\s+\d{4})([  ]*)$/);
    if (m && MESI.indexOf(m[2].toLowerCase()) >= 0) return { testo: m[1], spazi: m[3] };
    m = t.match(/(\S+)([  ]*)$/);
    if (!m) return null;
    var tok = m[1].replace(/^[(\[«"“']+/, "").replace(/^(€|euro|eur)/i, "");
    return tok ? { testo: tok, spazi: m[2] } : null;
  }

  /**
   * lettereDopo(testoPrima, opzioni) -> { daInserire, lettere, origine }
   * Calcola il testo da inserire al cursore, es. " (millecinquecento)".
   * Negli importi usa la forma con la barra se scelta nel pannello, altrimenti «virgola».
   */
  function lettereDopo(testoPrima, opzioni) {
    var fin = estraiFinale(testoPrima);
    if (!fin) throw new Error("Posiziona il cursore subito dopo un numero, una data o una quota.");
    var opz = {};
    for (var k in (opzioni || {})) opz[k] = opzioni[k];
    opz.tipo = "auto";
    opz.formatoNumero = opz.formatoNumero === "euro-barra" ? "euro-barra" : "lettere";
    var r;
    try { r = converti(fin.testo, opz); }
    catch (e) {
      if (opz.formatoNumero !== "euro-barra") {
        throw new Error("Prima del cursore c'è «" + fin.testo + "», che non riconosco come numero, data o quota.");
      }
      opz.formatoNumero = "lettere";
      try { r = converti(fin.testo, opz); }
      catch (e2) {
        throw new Error("Prima del cursore c'è «" + fin.testo + "», che non riconosco come numero, data o quota.");
      }
    }
    // "1.500|"  -> "1.500 (millecinquecento)|"
    // "1.500 |" -> "1.500 (millecinquecento) |"  (lo spazio resta prima della parola che segue)
    var daInserire = fin.spazi ? "(" + r.lettere + ")" + fin.spazi : " (" + r.lettere + ")";
    return { daInserire: daInserire, lettere: r.lettere, origine: fin.testo };
  }

  var api = {
    converti: converti,
    interoInLettere: interoInLettere,
    riconosciTipo: riconosciTipo,
    estraiFinale: estraiFinale,
    lettereDopo: lettereDopo
  };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.NumeriInLettere = api;
})(this);
