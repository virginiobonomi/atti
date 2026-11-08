# Lettore dei file RedAtto (.sap clausole, .sat minute): RTF con stop marcati da \sstag.
# Produce: {"p": [paragrafi], "stop": {n: definizione}}
# paragrafo: {"a": "j|c|l|r", "li": twip, "fi": twip, "x": [elementi]}
# elemento: {"t": testo, "b":1, "i":1, "u":1, "k":1}  oppure  {"s": n, "l": etichetta}  oppure  {"br": 1}
import re

SALTA = {"fonttbl", "colortbl", "stylesheet", "info", "header", "headerl", "headerr", "headerf",
         "footer", "footerl", "footerr", "footerf", "revtbl", "pict", "shp", "shpinst", "shprslt",
         "listtable", "listoverridetable", "fldinst", "picprop", "nonshppict", "shppict", "xmlnstbl",
         "rsidtbl", "generator", "themedata", "colorschememapping", "latentstyles", "datastore",
         "pn", "pntext", "bkmkstart", "bkmkend", "object", "objdata", "footnote", "annotation"}

TOKEN = re.compile(r"\\([a-zA-Z]+)(-?\d+)? ?|\\'([0-9a-fA-F]{2})|\\([^a-zA-Z])|([{}])|([^\\{}\r\n]+)|[\r\n]+")


def dec(s):
    return re.sub(r"\\'([0-9a-fA-F]{2})", lambda m: bytes([int(m.group(1), 16)]).decode("cp1252"), s)


def leggi(raw):
    if isinstance(raw, bytes):
        raw = raw.decode("cp1252", errors="replace")
    paragrafi, defs = [], {}
    stato = {"b": 0, "i": 0, "u": 0, "k": 0, "salta": False, "uc": 1}
    pila = []
    par = {"a": "l", "li": 0, "fi": 0}
    corrente = []           # elementi del paragrafo in corso
    stop_aperto = None      # [n, etichetta]
    chiudi_dopo = [None]    # n dello stop da chiudere dopo il prossimo carattere
    salta_car = [0]
    sstag = [None]
    f_anticipato = [None]          # testo accumulato in un gruppo \sstag

    def emetti(testo):
        if not testo:
            return
        nonlocal stop_aperto
        while testo:
            if stop_aperto is not None and len(stop_aperto) == 2:
                stop_aperto.append({k: 1 for k in ("b", "i", "u") if stato[k]})   # formattazione dello stop
            if chiudi_dopo[0] is not None and stop_aperto is not None:
                stop_aperto[1] += testo[0]
                el = {"s": stop_aperto[0], "l": stop_aperto[1]}
                el.update(stop_aperto[2])
                corrente.append(el)
                stop_aperto = None
                chiudi_dopo[0] = None
                testo = testo[1:]
                continue
            if stop_aperto is not None:
                stop_aperto[1] += testo
                return
            el = {"t": testo}
            for k in ("b", "i", "u", "k"):
                if stato[k]:
                    el[k] = 1
            if corrente and "t" in corrente[-1] and all(corrente[-1].get(k, 0) == el.get(k, 0) for k in ("b", "i", "u", "k")):
                corrente[-1]["t"] += testo
            else:
                corrente.append(el)
            return

    def fine_paragrafo():
        nonlocal corrente
        paragrafi.append({"a": par["a"], "li": par["li"], "fi": par["fi"], "x": corrente})
        corrente = []

    def gestisci_sstag(txt):
        nonlocal stop_aperto
        m = re.match(r"\s*([A-Za-z0-9_]+)\\sstagint0 ?(.*)$", txt, re.S)
        if not m:
            return
        nome, val = m.group(1), m.group(2)
        mi = re.match(r"I_FL_(\d+)$", nome)
        mf = re.match(r"F_FL_(\d+)$", nome)
        md = re.match(r"FL_(\d+)_([A-Z]+)$", nome)
        if mi:
            stop_aperto = [mi.group(1), ""]
            if f_anticipato[0] == mi.group(1):      # stop di un solo carattere: F precede I
                chiudi_dopo[0] = mi.group(1)
                f_anticipato[0] = None
        elif mf:
            if stop_aperto is not None:
                chiudi_dopo[0] = mf.group(1)
            else:
                f_anticipato[0] = mf.group(1)
        elif md:
            defs.setdefault(md.group(1), {})[md.group(2)] = dec(val).replace("\\\\", "\\").strip() if md.group(2) != "LV" else dec(val).replace("\\\\", "\\")

    for m in TOKEN.finditer(raw):
        parola, num, hexa, simb, graffa, testo = m.groups()
        if graffa == "{":
            pila.append(dict(stato))
            continue
        if graffa == "}":
            if sstag[0] is not None and stato.get("sst"):
                gestisci_sstag(sstag[0])
                sstag[0] = None
            if pila:
                vecchio = pila.pop()
                stato.clear(); stato.update(vecchio)
            continue
        if stato.get("sst"):
            # dentro \sstag: si accumula il testo grezzo, compresi i comandi
            sstag[0] += m.group(0) if not m.group(0).isspace() or m.group(0) == " " else ""
            continue
        if parola is not None:
            if parola == "sstag":
                stato["sst"] = True
                sstag[0] = ""
                continue
            if stato["salta"]:
                continue
            if stato.get("star"):
                stato["salta"] = True      # destinazione \* non gestita
                continue
            if salta_car[0] and parola:
                salta_car[0] = 0
            if parola in SALTA:
                stato["salta"] = True
                continue
            n = int(num) if num is not None else None
            if parola == "par":
                fine_paragrafo()
            elif parola == "pard":
                par.update({"a": "l", "li": 0, "fi": 0})
            elif parola in ("qj", "qc", "ql", "qr"):
                par["a"] = parola[1]
            elif parola == "li":
                par["li"] = n or 0
            elif parola == "fi":
                par["fi"] = n or 0
            elif parola == "plain":
                for k in ("b", "i", "u", "k"):
                    stato[k] = 0
            elif parola in ("b", "i", "scaps"):
                stato["k" if parola == "scaps" else parola] = 0 if n == 0 else 1
            elif parola in ("ul", "uld", "uldb", "ulw"):
                stato["u"] = 0 if n == 0 else 1
            elif parola == "ulnone":
                stato["u"] = 0
            elif parola == "line":
                if stop_aperto is None:
                    corrente.append({"br": 1})
            elif parola == "tab":
                emetti("\t")
            elif parola == "page":
                fine_paragrafo()
            elif parola == "uc":
                stato["uc"] = n or 1
            elif parola == "u":
                c = n if n >= 0 else n + 65536
                emetti(chr(c))
                salta_car[0] = stato["uc"]
            elif parola in ("rquote", "lquote"):
                emetti("'")
            elif parola in ("ldblquote", "rdblquote"):
                emetti('"')
            elif parola == "endash":
                emetti("–")
            elif parola == "emdash":
                emetti("—")
            elif parola == "bullet":
                emetti("•")
            continue
        if stato["salta"]:
            continue
        if hexa is not None:
            if salta_car[0]:
                salta_car[0] -= 1
                continue
            emetti(bytes([int(hexa, 16)]).decode("cp1252", errors="replace"))
            continue
        if simb is not None:
            if simb == "*":
                stato["star"] = True
                continue
            if simb == "~":
                emetti("\u00a0")
            elif simb in "\\{}":
                emetti(simb)
            elif simb == "_":
                emetti("-")
            continue
        if testo is not None:
            if stato.get("star") and not stato.get("sst"):
                # destinazione \* sconosciuta: si salta
                stato["salta"] = True
                continue
            if salta_car[0]:
                k = min(salta_car[0], len(testo))
                testo = testo[k:]
                salta_car[0] -= k
            emetti(testo)
    if corrente:
        fine_paragrafo()
    # togli i paragrafi vuoti finali
    while paragrafi and not paragrafi[-1]["x"]:
        paragrafi.pop()
    return {"p": paragrafi, "stop": defs}


def testo_semplice(doc):
    righe = []
    for p in doc["p"]:
        s = ""
        for x in p["x"]:
            if "t" in x:
                s += x["t"]
            elif "s" in x:
                s += "[" + x["l"] + "]"
            elif "br" in x:
                s += "\n"
        righe.append(s)
    return "\n".join(righe)


if __name__ == "__main__":
    import sys, json
    d = leggi(open(sys.argv[1], "rb").read())
    print(testo_semplice(d))
    if len(sys.argv) > 2:
        print(json.dumps(d["stop"], ensure_ascii=False, indent=1)[:3000])
