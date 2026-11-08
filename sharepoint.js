/*
 * Atti — collegamento al formulario dello studio su SharePoint (Microsoft Graph).
 * Accesso con l'account Microsoft 365 dell'utente (nested app authentication, MSAL).
 *
 * Cartella su SharePoint:
 *   Formulario Atti/formulario.json          formulario di base (convertito da RedAtto)
 *   Formulario Atti/Modifiche/*.json         una per ogni clausola o minuta modificata, nuova o eliminata
 * Le modifiche si applicano sopra il formulario di base. Ogni file conserva la cronologia delle versioni.
 */
(function (root) {
  "use strict";

  var CONFIG = {
    clientId: "34193037-52af-4a48-9ac8-a403d707a0ef",
    tenantId: "4bbd4e76-c604-4d93-afd8-f7bef1435555",
    host: "varesenotai.sharepoint.com",
    sito: "/sites/VareseNotai",
    cartella: "Formulario Atti",
    sottocartella: "Modifiche",
    scopes: ["Files.ReadWrite.All", "Sites.Read.All", "User.Read"]
  };
  var GRAPH = "https://graph.microsoft.com/v1.0";

  // ---------------------------------------------------------------- cache locale (IndexedDB)

  function db() {
    return new Promise(function (ok, ko) {
      if (!root.indexedDB) { ko(new Error("IndexedDB non disponibile")); return; }
      var rq = root.indexedDB.open("atti", 1);
      rq.onupgradeneeded = function () { rq.result.createObjectStore("dati"); };
      rq.onerror = function () { ko(rq.error); };
      rq.onsuccess = function () { ok(rq.result); };
    });
  }
  function cache(op, chiave, valore) {
    return db().then(function (d) {
      return new Promise(function (ok, ko) {
        var tx = d.transaction("dati", op === "get" || op === "keys" ? "readonly" : "readwrite");
        var st = tx.objectStore("dati");
        var r = op === "get" ? st.get(chiave) : op === "put" ? st.put(valore, chiave) : op === "del" ? st.delete(chiave) : st.getAllKeys();
        r.onsuccess = function () { ok(r.result); };
        r.onerror = function () { ko(r.error); };
      });
    });
  }

  // ---------------------------------------------------------------- accesso

  var pca = null, account = null;

  function msalPronto() {
    if (pca) return Promise.resolve(pca);
    if (!root.msal || !root.msal.createNestablePublicClientApplication) return Promise.reject(new Error("Libreria di accesso Microsoft non caricata."));
    return root.msal.createNestablePublicClientApplication({
      auth: { clientId: CONFIG.clientId, authority: "https://login.microsoftonline.com/" + CONFIG.tenantId },
      cache: { cacheLocation: "localStorage" }
    }).then(function (p) { pca = p; return p; });
  }

  function richiestaInterattiva(e) {
    return e && (e instanceof (root.msal && root.msal.InteractionRequiredAuthError || function () {}) ||
      /interaction_required|consent_required|login_required|no_account|no_tokens_found/i.test(String(e.errorCode || e.message || e)));
  }

  // interattivo = true solo dopo un clic dell'utente (la finestra di accesso ha bisogno di un gesto)
  function token(interattivo) {
    return msalPronto().then(function (p) {
      var conti = p.getAllAccounts ? p.getAllAccounts() : [];
      account = account || conti[0] || null;
      var req = { scopes: CONFIG.scopes };
      if (account) req.account = account;
      return p.acquireTokenSilent(req).catch(function (e) {
        if (!interattivo) {
          var err = new Error("Serve l'accesso con l'account dello studio.");
          err.accesso = true; err.originale = e;
          throw err;
        }
        return p.acquireTokenPopup({ scopes: CONFIG.scopes });
      });
    }).then(function (r) { if (r.account) account = r.account; return r.accessToken; });
  }

  // ---------------------------------------------------------------- Graph

  var interattivo = false;
  function graph(metodo, percorso, corpo, intestazioni, comeTesto) {
    return token(interattivo).then(function (t) {
      var h = { Authorization: "Bearer " + t };
      Object.keys(intestazioni || {}).forEach(function (k) { h[k] = intestazioni[k]; });
      var opz = { method: metodo, headers: h };
      if (corpo !== undefined) {
        opz.body = typeof corpo === "string" ? corpo : JSON.stringify(corpo);
        if (!h["Content-Type"]) h["Content-Type"] = "application/json";
      }
      return root.fetch(/^https:/.test(percorso) ? percorso : GRAPH + percorso, opz);
    }).then(function (r) {
      if (r.status === 204) return null;
      if (!r.ok) {
        return r.text().then(function (t) {
          var msg = t;
          try { msg = JSON.parse(t).error.message; } catch (e) { /* testo semplice */ }
          var err = new Error(msg || ("Errore " + r.status));
          err.status = r.status;
          throw err;
        });
      }
      return comeTesto ? r.text() : r.json();
    });
  }

  function seg(s) { return encodeURIComponent(s).replace(/%2F/g, "/"); }

  var driveId = null;
  function drive() {
    if (driveId) return Promise.resolve(driveId);
    return cache("get", "sp:drive").catch(function () { return null; }).then(function (c) {
      if (c) { driveId = c; return c; }
      return graph("GET", "/sites/" + CONFIG.host + ":" + CONFIG.sito + "?$select=id").then(function (s) {
        return graph("GET", "/sites/" + s.id + "/drive?$select=id");
      }).then(function (d) {
        driveId = d.id;
        return cache("put", "sp:drive", d.id).catch(function () {}).then(function () { return d.id; });
      });
    });
  }

  function percorsoCartella(sotto) {
    return "/drives/" + driveId + "/root:/" + seg(CONFIG.cartella + (sotto ? "/" + sotto : ""));
  }

  // ---------------------------------------------------------------- nomi dei file delle modifiche

  function hash(s) {
    var h = 2166136261;
    for (var i = 0; i < s.length; i++) { h ^= s.charCodeAt(i); h = Math.imul(h, 16777619); }
    return ("0000000" + (h >>> 0).toString(16)).slice(-8);
  }
  function nomeFile(sez, chiave) {
    var leggibile = chiave.replace(/\.(sap|sat)$/i, "").replace(/[\/\\]/g, " - ").replace(/["*:<>?|#%~&{}]/g, "_").slice(0, 90).trim();
    return (sez === "minute" ? "M " : "C ") + leggibile + " ~" + hash(sez + "|" + chiave) + ".json";
  }

  // ---------------------------------------------------------------- sincronizzazione

  function elencoFigli(percorso) {
    var tutti = [];
    function pagina(url) {
      return graph("GET", url).then(function (r) {
        tutti = tutti.concat(r.value || []);
        if (r["@odata.nextLink"]) return pagina(r["@odata.nextLink"]);
        return tutti;
      });
    }
    return pagina(percorso + ":/children?$select=id,name,eTag,lastModifiedDateTime,lastModifiedBy&$top=999").catch(function (e) {
      if (e.status === 404) return [];
      throw e;
    });
  }

  // restituisce { base: testo JSON, modifiche: [ {nome, id, eTag, dati} ], daCache: bool }
  function sincronizza(avanzamento) {
    avanzamento = avanzamento || function () {};
    var esito = { modifiche: [], daCache: false };
    return drive().then(function () {
      avanzamento("Controllo del formulario dello studio…");
      return graph("GET", percorsoCartella("formulario.json") + "?$select=id,eTag,lastModifiedDateTime");
    }).then(function (info) {
      return cache("get", "sp:base").catch(function () { return null; }).then(function (c) {
        if (c && c.eTag === info.eTag && c.testo) { esito.base = c.testo; esito.dataBase = info.lastModifiedDateTime; return; }
        avanzamento("Scaricamento del formulario (circa 10 MB)…");
        return graph("GET", "/drives/" + driveId + "/items/" + info.id + "/content", undefined, undefined, true).then(function (t) {
          esito.base = t; esito.dataBase = info.lastModifiedDateTime;
          return cache("put", "sp:base", { eTag: info.eTag, testo: t }).catch(function () {});
        });
      });
    }).then(function () {
      avanzamento("Controllo delle clausole modificate…");
      return elencoFigli(percorsoCartella(CONFIG.sottocartella));
    }).then(function (figli) {
      figli = figli.filter(function (f) { return /\.json$/i.test(f.name); });
      return cache("keys").catch(function () { return []; }).then(function (chiavi) {
        var presenti = {};
        figli.forEach(function (f) { presenti["sp:mod:" + f.id] = 1; });
        var togli = chiavi.filter(function (k) { return /^sp:mod:/.test(k) && !presenti[k]; });
        return Promise.all(togli.map(function (k) { return cache("del", k).catch(function () {}); }));
      }).then(function () {
        return Promise.all(figli.map(function (f) {
          return cache("get", "sp:mod:" + f.id).catch(function () { return null; }).then(function (c) {
            if (c && c.eTag === f.eTag) return c;
            return graph("GET", "/drives/" + driveId + "/items/" + f.id + "/content", undefined, undefined, true).then(function (t) {
              var v = { nome: f.name, id: f.id, eTag: f.eTag, dati: JSON.parse(t) };
              return cache("put", "sp:mod:" + f.id, v).catch(function () {}).then(function () { return v; });
            });
          });
        }));
      });
    }).then(function (mods) {
      esito.modifiche = mods.filter(Boolean);
      return esito;
    });
  }

  // in mancanza di rete: ultimo formulario scaricato
  function daCache() {
    return cache("get", "sp:base").then(function (c) {
      if (!c || !c.testo) return null;
      return cache("keys").then(function (chiavi) {
        return Promise.all(chiavi.filter(function (k) { return /^sp:mod:/.test(k); }).map(function (k) { return cache("get", k); }));
      }).then(function (mods) { return { base: c.testo, modifiche: mods.filter(Boolean), daCache: true }; });
    }).catch(function () { return null; });
  }

  // applica le modifiche sopra il formulario di base
  function applica(F, modifiche) {
    var info = {};
    (modifiche || []).forEach(function (m) {
      var d = m.dati;
      if (!d || !d.sez || !d.chiave || (d.sez !== "clausole" && d.sez !== "minute")) return;
      var coll = F[d.sez];
      var cat = (d.doc && d.doc.cat !== undefined ? d.doc.cat : (coll[d.chiave] ? coll[d.chiave].cat : "")) || "";
      var catKey = cat.toLowerCase();
      if (d.sez === "clausole") {
        Object.keys(F.categorie).forEach(function (c) {
          F.categorie[c] = F.categorie[c].filter(function (k) { return k !== d.chiave; });
        });
      }
      if (d.doc) {
        coll[d.chiave] = d.doc;
        if (d.sez === "clausole") (F.categorie[catKey] = F.categorie[catKey] || []).push(d.chiave);
      } else delete coll[d.chiave];
      info[d.sez + "|" + d.chiave] = { id: m.id, eTag: m.eTag, nome: m.nome, autore: d.autore, data: d.data, eliminata: !d.doc };
    });
    return info;
  }

  function assicuraCartellaModifiche() {
    return graph("GET", percorsoCartella(CONFIG.sottocartella) + "?$select=id").catch(function (e) {
      if (e.status !== 404) throw e;
      return graph("GET", percorsoCartella("") + "?$select=id").then(function (padre) {
        return graph("POST", "/drives/" + driveId + "/items/" + padre.id + "/children",
          { name: CONFIG.sottocartella, folder: {}, "@microsoft.graph.conflictBehavior": "replace" });
      });
    });
  }

  var me = null;
  function utente() {
    if (me) return Promise.resolve(me);
    return graph("GET", "/me?$select=displayName,mail").then(function (u) { me = u.displayName || u.mail || ""; return me; }).catch(function () { return ""; });
  }

  // salva una clausola o minuta (doc = null per eliminarla). eTag: quello letto, per non sovrascrivere modifiche altrui
  function salva(sez, chiave, doc, eTagPrecedente) {
    var nome = nomeFile(sez, chiave);
    return drive().then(assicuraCartellaModifiche).then(utente).then(function (autore) {
      var dati = { sez: sez, chiave: chiave, doc: doc, autore: autore, data: new Date().toISOString() };
      var h = { "Content-Type": "application/json" };
      var url = percorsoCartella(CONFIG.sottocartella + "/" + nome) + ":/content";
      if (eTagPrecedente) h["If-Match"] = eTagPrecedente;
      else url += "?@microsoft.graph.conflictBehavior=fail";
      return graph("PUT", url, JSON.stringify(dati, null, 1), h).then(function (item) {
        var v = { nome: item.name, id: item.id, eTag: item.eTag, dati: dati };
        return cache("put", "sp:mod:" + item.id, v).catch(function () {}).then(function () { return v; });
      });
    }).catch(function (e) {
      if (e.status === 412 || e.status === 409) {
        var err = new Error("Nel frattempo qualcun altro ha modificato questa voce: chiudi e riapri il pannello per vedere la versione aggiornata.");
        err.conflitto = true;
        throw err;
      }
      throw e;
    });
  }

  // toglie la modifica: torna il testo del formulario di base
  function ripristina(idFile) {
    return drive().then(function () {
      return graph("DELETE", "/drives/" + driveId + "/items/" + idFile);
    }).then(function () { return cache("del", "sp:mod:" + idFile).catch(function () {}); });
  }

  root.Archivio = {
    CONFIG: CONFIG, sincronizza: sincronizza, daCache: daCache, applica: applica, salva: salva, ripristina: ripristina,
    accedi: function () { interattivo = true; return token(true); },
    nomeFile: nomeFile
  };
})(this);
