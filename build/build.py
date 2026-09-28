#!/usr/bin/env python3
"""
Generează pachetele PDF pentru fiecare oră din Google Calendar (9E și 11P1)
și datele pentru site-ul beldugan.github.io/lectii.

Intrări:
  --src     folderul „Planificari si Manuale 2026 2027”
  --events  exportul JSON al evenimentelor din Google Calendar
  --out     folderul site-ului (repository-ul lectii)
  --password parola pentru PDF-urile cu teste și bareme

Rulare:  python3 build.py --src ... --events events.json --out ../site --password XXXX
"""
import argparse, hashlib, json, os, re, subprocess, sys, tempfile, unicodedata
from collections import defaultdict
from datetime import datetime

from pypdf import PdfReader
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

pdfmetrics.registerFont(TTFont("DV", "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"))
pdfmetrics.registerFont(TTFont("DVB", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"))

ROMAN = {"I": 1, "II": 2, "III": 3, "IV": 4}

# ---------------------------------------------------------------- surse
MOD = {
    ("9e", 1): dict(name="Modulul I – Bazele desenului tehnic",
                    plan="9E/Planuri lectie/Proiecte_lectie_Modul_I_Desen_tehnic.pdf",
                    man="9E/Manuale/Manual_Modul_I_Bazele_desenului_tehnic_IX_T-LT_{ed}.pdf"),
    ("9e", 2): dict(name="Modulul II – Lăcătușerie generală",
                    plan="9E/Planuri lectie/Proiecte_lectie_Modul_II_Lacatuserie_generala.pdf",
                    man="9E/Manuale/Manual_Modul_II_Lacatusarie_generala_IX_T-LT-IP_{ed}.pdf"),
    ("9e", 3): dict(name="Modulul III – Organe de mașini",
                    plan="9E/Planuri lectie/Proiecte_lectie_Modul_III_Organe_de_masini.pdf",
                    man="9E/Manuale/Manual_Modul_III_Organe_de_masini_IX_T-LT_{ed}.pdf"),
    ("9e", 4): dict(name="Modulul IV – Asamblări mecanice",
                    plan="9E/Planuri lectie/Proiecte_lectie_Modul_IV_Asamblari_mecanice.pdf",
                    man="9E/Manuale/Manual_Modul_IV_Asamblari_mecanice_IX_T-IP_{ed}.pdf"),
    ("11p1", 1): dict(name="Modulul I – Întreținerea mașinilor, utilajelor și instalațiilor",
                      plan="11P1/Planuri lectie/Proiecte_lectie_Modul_I_Intretinerea_MUI_XI_P1_2026-2027.pdf",
                      man="11P1/Manuale/Manual_Modul_I_Intretinerea_masinilor_utilajelor_si_instalatiilor_XI_P1_LT-IP_{ed}.pdf"),
    ("11p1", 2): dict(name="Modulul II – Repararea subansamblurilor mașinilor, utilajelor și instalațiilor",
                      plan="11P1/Planuri lectie/Proiecte_lectie_Modul_II_Repararea_subansamblurilor_XI_P1_2026-2027.pdf",
                      man="11P1/Manuale/Manual_Modul_II_Repararea_subansamblurilor_MUI_XI_P1_LT-IP_{ed}.pdf"),
    ("11p1", 3): dict(name="Modulul III – Instalații de ridicat și transportat",
                      plan="11P1/Planuri lectie/Proiecte_lectie_Modul_III_Instalatii_ridicat_transportat_XI_P1_2026-2027.pdf",
                      man="11P1/Manuale/Manual_Modul_III_Instalatii_de_ridicat_si_transportat_XI_P1_LT-IP_{ed}.pdf"),
}
CAIET_9E = "9E/Caiet de practica/Caiet_practica_comasata_IX-E_2026-2027_{ed}.pdf"
CAIETE_11 = {
    "CELCO": "11P1/Caiete de practica/Caiet_practica_stagiu_CDL_XI_P1_2026-2027_CELCO_{ed}.pdf",
    "COMVEX": "11P1/Caiete de practica/Caiet_practica_stagiu_CDL_XI_P1_2026-2027_COMVEX_{ed}.pdf",
    "DP World": "11P1/Caiete de practica/Caiet_practica_stagiu_CDL_XI_P1_2026-2027_DP_World_CSCT_{ed}.pdf",
}
TYPE_NAME = {"t": "ora de teorie (T)", "lt": "ora de laborator tehnologic (LT)",
             "ip": "ora de instruire practică (IP)", "stagiu": "stagiu de pregătire practică",
             "caip": "CAIP – practică comasată"}
CLS_NAME = {"9e": "9E", "11p1": "11P1"}


# ---------------------------------------------------------------- documente
class Doc:
    cache = {}

    @classmethod
    def get(cls, src, rel):
        path = os.path.join(src, rel)
        if path not in cls.cache:
            cls.cache[path] = Doc(path, rel)
        return cls.cache[path]

    def __init__(self, path, rel):
        self.path, self.rel = path, rel
        txt = subprocess.run(["pdftotext", "-layout", path, "-"], capture_output=True, text=True).stdout
        self.pages = txt.split("\f")
        if self.pages and not self.pages[-1].strip():
            self.pages.pop()
        self.n = len(PdfReader(path).pages)
        self.toc = self._toc()

    def _toc(self):
        """Cuprinsul: listă (titlu, pagină_start, pagină_sfârșit), cu titlurile rupte pe două rânduri lipite."""
        entries, in_toc = [], False
        for pi, pg in enumerate(self.pages[:12]):
            lines = [l for l in pg.splitlines() if l.strip()]
            has = [re.match(r"\s*(.*?)\s*\.{3,}\s*(\d+)\s*$", l) for l in lines]
            if not any(has):
                if in_toc:
                    break
                continue
            in_toc, pending = True, ""
            for l, m in zip(lines[1:], has[1:]):          # primul rând = antetul paginii
                if l.strip() in ("Cuprins",):
                    continue
                if m:
                    t = re.sub(r"\s+", " ", (pending + " " + m.group(1)).strip(" ."))
                    # rânduri de cuprins lipite: „… 56 Lucrarea 17. …”
                    while True:
                        mm = re.match(r"^(.*?)[\s.]+(\d+)\s+((?:Lucrarea \d+\.|S\d+ ·|E\d+\.|Barem |Fișa |Anexa |"
                                      r"Capitolul |Unitatea |Săptămâna ).*)$", t)
                        if not mm:
                            break
                        entries.append([mm.group(1).strip(" ."), int(mm.group(2))])
                        t = mm.group(3)
                    entries.append([t, int(m.group(2))])
                    pending = ""
                else:
                    pending = (pending + " " + l.strip()).strip()
        entries.sort(key=lambda e: e[1])
        # unde începe efectiv fiecare titlu (sus pe pagină sau la mijloc)
        norm = lambda s: re.sub(r"[^\w]+", "", s).lower()
        pages_norm = [[norm(l) for l in pg.splitlines() if l.strip()] for pg in self.pages]
        top = []
        for t, p in entries:
            key = norm(" ".join(t.split("·")[-1].split()[:5]))[:22] or norm(t)[:22]
            at_top = True
            if 1 <= p <= len(pages_norm):
                lines = pages_norm[p - 1][1:]          # fără antet
                idx = next((i for i, l in enumerate(lines) if key and key in l), None)
                if idx is not None and idx > 4:
                    at_top = False
            top.append(at_top)
        out = []
        for i, (t, p) in enumerate(entries):
            end = self.n
            for j in range(i + 1, len(entries)):
                q = entries[j][1]
                if q > p:
                    is_eval = re.match(r"(E\d+|PE|Barem|F\d+)\b", entries[j][0])
                    end = q - 1 if (top[j] or is_eval) else q
                    break
            out.append((t, p, max(p, min(end, self.n))))
        return out

    def section_of_page(self, p, pattern):
        hits = [(t, a, b) for t, a, b in self.toc if a <= p <= b and re.match(pattern, t)]
        return max(hits, key=lambda h: h[1]) if hits else None

    def sections(self, pattern):
        return [(t, a, b) for t, a, b in self.toc if re.match(pattern, t)]


def plan_ranges(doc):
    """Proiectele de lecție: nr -> (S, pagina_start, pagina_sfârșit)."""
    starts = []
    for i, pg in enumerate(doc.pages):
        for m in re.finditer(r"PROIECT DE LECȚIE nr\.\s*(\d+)\s*·\s*S(\d+)", pg):
            starts.append((int(m.group(1)), int(m.group(2)), i + 1))
    out = {}
    for k, (nr, s, p) in enumerate(starts):
        end = starts[k + 1][2] - 1 if k + 1 < len(starts) else doc.n
        out[nr] = (s, p, end)
    return out


# ---------------------------------------------------------------- evenimente
def parse_event(ev):
    s, d = ev.get("summary", ""), ev.get("description", "") or ""
    st = ev["start"].get("dateTime")
    if not st:
        return None
    dt = datetime.fromisoformat(st)
    m = re.match(r"^\[(T|LT|IP)\] (11P1 )?M (IV|III|II|I) –\s*(.*)", s)
    if m:
        cls = "11p1" if m.group(2) else "9e"
        typ, mod, title = m.group(1).lower(), ROMAN[m.group(3)], m.group(4)
    elif s.startswith("[Stagiu CDL] 11P1"):
        cls, typ, mod, title = "11p1", "stagiu", None, s.split("–", 1)[-1].strip()
    elif s.startswith("[Stagiu CDEOȘ] 9E"):
        cls, typ, mod, title = "9e", "stagiu", None, s.split("–", 2)[-1].strip()
    elif s.startswith("[CAIP] 9E"):
        cls, typ, mod, title = "9e", "caip", None, s.split("–", 1)[-1].strip()
    else:
        return None
    link = re.search(r"lectii/o\.html\?o=([\w-]+)", d)
    if link:
        oid = link.group(1)
    else:
        suf = f"m{mod}-{typ}" if mod else typ
        oid = f"{cls}-{dt:%Y-%m-%d-%H%M}-{suf}"
    wk = re.search(r"\bS(\d+)\b", d) or re.search(r"\bS(\d+)\b", s)
    nr = re.search(r"Proiect de lecție nr\.\s*(\d+)", d)
    pages = []
    for line in d.splitlines():
        if line.strip().startswith("Manual"):
            for a, b in re.findall(r"(\d+)\s*(?:[–-]\s*(\d+))?", line.split("p.", 1)[-1] if "p." in line else ""):
                pages += list(range(int(a), int(b or a) + 1))
    ev_line = next((l for l in d.splitlines() if l.startswith("Evaluare")), "")
    evals = sorted(set(re.findall(r"\b(E\d+|PE)\b", ev_line)))
    dayn = re.search(r"[Zz]iua (\d+)", d)
    return dict(id=oid, cls=cls, type=typ, mod=mod, title=title, summary=s, desc=d,
                date=dt.strftime("%Y-%m-%d"), time=dt.strftime("%H:%M"),
                end=datetime.fromisoformat(ev["end"]["dateTime"]).strftime("%H:%M"),
                week=int(wk.group(1)) if wk else None, nr=int(nr.group(1)) if nr else None,
                pages=pages, evals=evals, day=int(dayn.group(1)) if dayn else None,
                has_link=bool(link), event_id=ev["id"])


# ---------------------------------------------------------------- pachete
def resolve(e, src, warn):
    """Întoarce (parts_public, parts_locked); o parte = (eticheta, doc, a, b)."""
    pub, lock = [], []
    if e["mod"]:
        info = MOD[(e["cls"], e["mod"])]
        plan = Doc.get(src, info["plan"])
        pr = plan_ranges(plan)
        # săptămâna din calendar are prioritate; numărul proiectului doar dacă săptămâna nu are proiect
        nr = next((k for k, (s, a, b) in pr.items() if s == e["week"]), None)
        if nr is None:
            nr = e["nr"]
        if nr in pr:
            s, a, b = pr[nr]
            pub.append((f"Proiect de lecție nr. {nr} (S{s})", plan, a, b))
        else:
            warn(e, f"fără proiect de lecție pentru S{e['week']}")
        man = Doc.get(src, info["man"].format(ed="ELEV"))
        secs = man.sections(rf"S{e['week']} ·") if e["week"] else []
        if not secs:                       # săptămână fără secțiune proprie: după paginile din calendar
            for p in e["pages"]:
                sec = man.section_of_page(p, r"S\d+ ·")
                if sec and sec not in secs:
                    secs.append(sec)
        if not secs:
            warn(e, "fără pagini din manual")
        for t, a, b in secs:
            pub.append((f"Manual (ed. elev) – {t}", man, a, b))
        if e["evals"]:
            prof = Doc.get(src, info["man"].format(ed="PROFESOR"))
            for code in e["evals"]:
                found = prof.sections(rf"{code}\.") + prof.sections(rf"Barem {code}\b")
                if not found:
                    warn(e, f"nu găsesc {code} în manualul profesorului")
                for t, a, b in found:
                    lock.append((f"Manual (ed. profesor) – {t}", prof, a, b))
    elif e["cls"] == "11p1":          # stagiu CDL
        for op, rel in CAIETE_11.items():
            m = re.search(rf"^{re.escape(op)}: Lucrarea (\d+)", e["desc"], re.M)
            if not m:
                continue
            n = int(m.group(1))
            el = Doc.get(src, rel.format(ed="ELEV"))
            secs = el.sections(rf"Lucrarea {n}\.")
            if not secs:
                warn(e, f"{op}: nu găsesc lucrarea {n}")
            for t, a, b in secs:
                pub.append((f"Caiet de practică {op} (ed. elev) – {t}", el, a, b))
            if re.search(r"[Ee]valuare", " ".join(t for t, _, _ in secs)):
                pf = Doc.get(src, rel.format(ed="PROFESOR"))
                for t, a, b in pf.sections(rf"Lucrarea {n}\."):
                    lock.append((f"Caiet de practică {op} (ed. profesor) – {t}", pf, a, b))
    else:                               # 9E: stagiu CDEOȘ sau CAIP
        el = Doc.get(src, CAIET_9E.format(ed="ELEV"))
        cl = next((l for l in e["desc"].splitlines() if l.startswith("Caiet")), "")
        nums = [int(x) for x in re.findall(r"lucrarea (\d+)", cl)]
        if e["type"] == "caip" and e["day"]:
            nums = [25 + e["day"]]
        for n in nums:
            secs = el.sections(rf"Lucrarea {n}\.")
            if not secs:
                warn(e, f"nu găsesc lucrarea {n}")
            for t, a, b in secs:
                pub.append((f"Caiet de practică (ed. elev) – {t}", el, a, b))
                if re.search(r"[Ee]valuare", t):
                    pf = Doc.get(src, CAIET_9E.format(ed="PROFESOR"))
                    for t2, a2, b2 in pf.sections(rf"Lucrarea {n}\."):
                        lock.append((f"Caiet de practică (ed. profesor) – {t2}", pf, a2, b2))
    return pub, lock


def cap(s):
    return s[:1].upper() + s[1:]


def wrap(c, text, x, y, width, font, size, lead):
    words, line = text.split(), ""
    for w in words:
        if pdfmetrics.stringWidth((line + " " + w).strip(), font, size) > width:
            c.drawString(x, y, line); y -= lead; line = w
        else:
            line = (line + " " + w).strip()
    if line:
        c.drawString(x, y, line); y -= lead
    return y


def cover(path, e, parts, locked=False):
    W, H = A4
    c = canvas.Canvas(path, pagesize=A4)
    c.setTitle(e["title"])
    c.setFillColorRGB(0.12, 0.23, 0.42); c.rect(0, H - 110, W, 110, stroke=0, fill=1)
    c.setFillColorRGB(1, 1, 1)
    c.setFont("DV", 10); c.drawString(50, H - 40, "LICEUL TEHNOLOGIC „LAZĂR EDELEANU” NĂVODARI · Catedra Mecanică · 2026–2027")
    c.setFont("DVB", 20)
    head = f"Clasa {CLS_NAME[e['cls']]}" + (f" · S{e['week']}" if e["week"] else "")
    c.drawString(50, H - 75, head + ("  ·  TEST / BAREM" if locked else ""))
    c.setFillColorRGB(0.1, 0.1, 0.1)
    y = H - 150
    mod = MOD[(e["cls"], e["mod"])]["name"] if e["mod"] else cap(TYPE_NAME[e["type"]])
    c.setFont("DV", 12); y = wrap(c, mod, 50, y, W - 100, "DV", 12, 16)
    if e["mod"]:
        c.drawString(50, y, cap(TYPE_NAME[e["type"]])); y -= 16
    y -= 10
    c.setFont("DVB", 16); y = wrap(c, e["title"], 50, y, W - 100, "DVB", 16, 21)
    y -= 20
    c.setFont("DVB", 11); c.drawString(50, y, "Conținutul pachetului"); y -= 18
    c.setFont("DV", 10.5)
    pg = 2
    for label, doc, a, b in parts:
        n = b - a + 1
        txt = f"p. {pg}–{pg + n - 1}   {label}   (sursa: p. {a}–{b})" if n > 1 else f"p. {pg}   {label}   (sursa: p. {a})"
        y = wrap(c, txt, 60, y, W - 120, "DV", 10.5, 14); y -= 4
        pg += n
    if e["evals"] and not locked:
        y -= 10
        c.setFont("DV", 10.5)
        y = wrap(c, f"Evaluare în această oră: {', '.join(e['evals'])}. Testul și baremul sunt în PDF-ul separat, protejat cu parolă.",
                 50, y, W - 100, "DV", 10.5, 14)
    c.setFont("DV", 8.5); c.setFillColorRGB(0.4, 0.4, 0.4)
    c.drawString(50, 40, "prof. ing. Adrian-Mircea Beldugan · beldugan.github.io/lectii")
    c.showPage(); c.save()


def assemble(out, e, parts, locked, password, tmp):
    cv = os.path.join(tmp, "cover.pdf")
    cover(cv, e, parts, locked)
    args = ["qpdf", "--empty", "--pages", cv, "1"]
    for _, doc, a, b in parts:
        args += [doc.path, f"{a}-{b}"]
    args += ["--", "--object-streams=generate", "--compress-streams=y", "--recompress-flate",
             "--remove-unreferenced-resources=yes"]
    if locked:
        args += ["--encrypt", password, password + "-owner", "256", "--"]
    args.append(out)
    r = subprocess.run(args, capture_output=True, text=True)
    if r.returncode not in (0, 3):
        raise RuntimeError(r.stderr)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--src", required=True)
    ap.add_argument("--events", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--password", required=True)
    ap.add_argument("--only-week", type=int)
    A = ap.parse_args()

    evs = [x for x in (parse_event(ev) for ev in json.load(open(A.events))) if x]
    evs.sort(key=lambda e: (e["date"], e["time"]))
    warnings = []
    warn = lambda e, msg: warnings.append(f"{e['id']}: {msg}")

    built, index = {}, []
    os.makedirs(os.path.join(A.out, "pdf"), exist_ok=True)
    with tempfile.TemporaryDirectory() as tmp:
        for e in evs:
            pub, lock = resolve(e, A.src, warn)
            rec = {k: e[k] for k in ("id", "cls", "type", "mod", "title", "date", "time", "end", "week", "evals")}
            rec["module"] = MOD[(e["cls"], e["mod"])]["name"] if e["mod"] else TYPE_NAME[e["type"]]
            rec["parts"] = [lbl for lbl, *_ in pub]
            for kind, parts, locked in (("pdf", pub, False), ("test", lock, True)):
                if not parts:
                    continue
                key = json.dumps([e["cls"], e["mod"], e["type"], e["week"], e["title"], locked,
                                  [(d.rel, a, b) for _, d, a, b in parts]])
                if key not in built:
                    h = hashlib.sha1(key.encode()).hexdigest()[:6]
                    base = f"{e['cls']}-" + (f"m{e['mod']}-" if e["mod"] else "") + \
                           (f"s{e['week']:02d}-" if e["week"] else "") + e["type"] + \
                           ("-test" if locked else "") + f"-{h}.pdf"
                    rel = f"pdf/{base}"
                    if A.only_week is None or e["week"] == A.only_week:
                        assemble(os.path.join(A.out, rel), e, parts, locked, A.password, tmp)
                    built[key] = rel
                rec[kind] = built[key]
            index.append(rec)
    os.makedirs(os.path.join(A.out, "data"), exist_ok=True)
    json.dump(index, open(os.path.join(A.out, "data", "ore.json"), "w"), ensure_ascii=False, separators=(",", ":"))
    missing = [dict(id=e["id"], event_id=e["event_id"]) for e in evs if not e["has_link"]]
    json.dump(missing, open(os.path.join(A.out, "..", "missing_links.json"), "w"), indent=1)
    print(f"{len(index)} ore, {len(set(built.values()))} PDF-uri, {len(missing)} fără link în calendar")
    print(f"{len(warnings)} avertismente")
    for w in warnings:
        print("  ", w)


if __name__ == "__main__":
    main()
