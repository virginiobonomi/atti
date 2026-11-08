# Converte l'intero formulario RedAtto (cartelle Clausole e Minute Atti) in formulario.json
import os, sys, json, re, unicodedata, time
sys.path.insert(0, os.path.dirname(__file__))
import rtfred

def nfc(s):
    return unicodedata.normalize("NFC", s)

def chiave(rel):
    return nfc(rel.replace("\\", "/").strip("/")).lower()

def compatta_stop(d):
    fo = d.get("FO", "")
    m = re.match(r"([A-Z]+?)(\d+)S_([A-Z0-9]+)$", fo)
    out = {"n": d.get("ID", "").strip()}
    if m:
        out["t"], out["f"], out["k"] = m.group(1), int(m.group(2)), m.group(3)
    else:
        out["t"], out["f"], out["k"] = "T", 1, "TE"
    if d.get("FT") and d["FT"] != "Standard":
        out["ft"] = d["FT"]
    if d.get("LV"):
        if out["t"] == "LC":
            voci = []
            for it in d["LV"].split("~"):
                cat, _, f = it.partition("|")
                voci.append([chiave(cat), "*" if f.strip() == "*" else chiave(f)])
            out["v"] = voci
        else:
            out["v"] = d["LV"].split("~")
    if d.get("VE") == "True":
        out["ve"] = 1
    if d.get("SV") == "True":
        out["sv"] = 1
    for k, kk in (("PR", "pr"), ("PO", "po"), ("VA", "va")):
        if d.get(k):
            out[kk] = d[k]
    return out

def converti(path):
    doc = rtfred.leggi(open(path, "rb").read())
    usati = {x["s"] for p in doc["p"] for x in p["x"] if "s" in x}
    stop = {}
    for n, d in doc["stop"].items():
        if n not in usati:
            continue
        stop[n] = compatta_stop(d)
        co = re.sub(r"^FL_", "", d.get("CO", ""))
        if co and co != n and co in usati:
            stop[n]["co"] = co          # variabile copia di un altro stop
    par = []
    for p in doc["p"]:
        q = {"x": p["x"]}
        if p["a"] != "l": q["a"] = p["a"]
        if p["li"]: q["li"] = p["li"]
        if p["fi"]: q["fi"] = p["fi"]
        par.append(q)
    return {"p": par, "s": stop}

def main(radice, uscita):
    t = time.time()
    out = {"versione": time.strftime("%Y-%m-%d %H:%M"), "clausole": {}, "categorie": {}, "minute": {}}
    for sezione, dove in (("clausole", "Clausole"), ("minute", "Minute Atti")):
        base = os.path.join(radice, dove)
        for r, ds, fs in os.walk(base):
            ds.sort(); fs.sort()
            for f in fs:
                if not f.lower().endswith((".sap", ".sat")):
                    continue
                rel = os.path.relpath(os.path.join(r, f), base)
                k = chiave(rel)
                d = converti(os.path.join(r, f))
                d["nome"] = nfc(os.path.splitext(f)[0])
                d["cat"] = nfc(os.path.dirname(rel))
                out[sezione][k] = d
                if sezione == "clausole":
                    out["categorie"].setdefault(chiave(os.path.dirname(rel)), []).append(k)
    json.dump(out, open(uscita, "w"), ensure_ascii=False, separators=(",", ":"))
    print(len(out["clausole"]), "clausole,", len(out["minute"]), "minute,", len(out["categorie"]), "categorie,",
          os.path.getsize(uscita) // 1024, "KB,", round(time.time() - t, 1), "s")

if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
