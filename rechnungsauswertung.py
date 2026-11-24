"""Rechnungsauswertung - fasst die Beträge von Lieferantenrechnungen pro Kommission zusammen.

Lieferanten: siehe PROFILE weiter unten.

Bedienung:
  - PDF(s) auf die .exe ziehen, oder
  - .exe doppelklicken und PDF(s) auswählen.
Für jede Rechnung wird neben dem PDF eine Excel-Datei "<PDF-Name>_Kommissionen.xlsx" erzeugt.
"""
import os
import re
import sys
from collections import OrderedDict
from decimal import Decimal, ROUND_DOWN

import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

BETRAG = r"(-?[\d.]+,\d{2})"
OHNE = "OHNE KOMMISSION"


def zahl(s):
    return Decimal(s.replace(".", "").replace(",", "."))


def lies_text(pfad):
    with pdfplumber.open(pfad) as pdf:
        return "\n".join(seite.extract_text() or "" for seite in pdf.pages)


# ---------------------------------------------------------------------------
# Lieferanten-Profile: pro Lieferant die Leseregeln (reguläre Ausdrücke).
#   erkennung  Text, der nur auf Rechnungen dieses Lieferanten vorkommt
#   rechnung   Rechnungsnummer (Gruppe 1) und Datum (Gruppe 2); oder nummer/datum einzeln
#   reset      Zeile, ab der eine neue Lieferung beginnt (Kommission wird vergessen)
#   kommission Zeile mit der Kommission (Gruppe 1 = Text nach der Markierung)
#   betrag     Zeile, deren Betrag (Gruppe 1) zur aktuellen Kommission zählt
#   summe_pos  Summe aller Positionen laut PDF (Kontrolle, optional)
#   zuschlag   Zuschlagszeile im Rechnungsfuß (Gruppe 1 = Name, Gruppe 2 = Betrag)
#   netto      Nettosumme laut PDF (Gruppe 1; mehrere Treffer werden addiert)
#   ende       ab dieser Zeile wird nichts mehr gelesen
# Neuer Lieferant = neuer Eintrag hier.
# ---------------------------------------------------------------------------
PROFILE = [
    {
        "name": "Schachermayer",
        "erkennung": r"Schachermayer GmbH",
        "rechnung": r"Rechnungsnummer:\s*(\d+)\s+Datum:\s*([\d.]+)",
        "reset": r"^Auftrag:\s*\d+",
        "kommission": r"^(?:Best\.Nr\.|Kommission):\s*(.*)",
        "betrag": r"^Zwischensumme Auftrag:.*?" + BETRAG + r"$",
        "summe_pos": r"^Summe Positionen\s+" + BETRAG,
        "zuschlag": r"^(.+?)\s+" + BETRAG + r"$",
        "netto": r"^Mehrwertsteuer.*?aus\s+" + BETRAG,
        "ende": r"^Endbetrag",
    },
    {
        "name": "heroal",
        "erkennung": r"heroal – Johann Henkenjohann",
        "rechnung": r"(?m)^\d\d\.\d{3} (\d{3}\.\d{3}) (\d\d?\.\d\d\.\d\d)$",
        "reset": r"^\$ Auftrag:",
        "kommission": r"^Kom\.:\s*(.*)",
        "betrag": r"^Zwischen-Summe\s+" + BETRAG + r"$",
        "netto": r"^Rechnungsbetrag\s*:\s*EUR\s+" + BETRAG,
        "ende": r"^Rechnungsbetrag",
    },
    {
        "name": "Vetrotech",
        "erkennung": r"VETROTECH SAINT-GOBAIN",
        "nummer": r"(?m)^Nummer\s+(\d+)",
        "datum": r"(?m)^Datum\s+([\d.]+)",
        "kommission": r"^Kommission:\s*(.*)",
        "betrag": r"^\d+ \S+ \d+ \d+ x \d+ .*?\s" + BETRAG + r"$",
        "summe_pos": r"Zwischensumme EUR\s+" + BETRAG,
        "zuschlag": r"(\w*(?:zuschlag|pauschale|zustellung|entladung)\w*)\b.*?\s" + BETRAG + r"$",
        "netto": r"\bNetto\s+" + BETRAG + r"$",
        "ende": r"Gesamtsumme EUR",
    },
    {
        "name": "Süd-Metall",
        "erkennung": r"Süd-Metall Beschläge GmbH",
        "nummer": r"Rechnungsnr\.\s+(\S+)",
        "datum": r"Datum:\s+([\d.]+)",
        "kommission": r"^(?i:kom)\.?:?\s+(\S.*)",
        "betrag": r"^\d\d\.\d\d\.\d{4} .*\s" + BETRAG + r"$",
        "netto": r"^Total EUR\s+" + BETRAG + r"$",
        "ende": r"^Total EUR",
    },
]


def finde_profil(text):
    for p in PROFILE:
        if re.search(p["erkennung"], text):
            return p
    return None


def kommission_key(text):
    # 5-stellige Nummer am Anfang; sonst die erste alleinstehende 5-stellige Nummer im Text
    m = re.match(r"\s*(\d{5})(?!\d)", text or "") or re.search(r"(?<!\d)(\d{5})(?!\d)", text or "")
    return m.group(1) if m else OHNE


def beschreibung(text):
    # "26010_RK SCHAFFERER WA LANS" / "26010/MKO X" -> "SCHAFFERER WA LANS" / "X"
    if kommission_key(text) == OHNE:
        return (text or "").strip()
    if not re.match(r"\s*\d{5}(?!\d)", text):
        return text.strip()
    return re.sub(r"^\s*\d{5}(?:[_/]?[A-Za-z]{2,4}(?=\s|$))?\s*[-–]?\s*", "", text).strip()


def auswerten(text, p):
    info = {"nummer": "", "datum": "", "lieferant": p["name"]}
    if "rechnung" in p:
        if m := re.search(p["rechnung"], text):
            info["nummer"], info["datum"] = m.group(1), m.group(2)
    else:
        for feld in ("nummer", "datum"):
            if m := re.search(p[feld], text):
                info[feld] = m.group(1)

    kommissionen = OrderedDict()  # key -> {"namen": [], "summe": Decimal}
    zuschlaege = []               # (Bezeichnung, Betrag)
    summe_pos = None
    netto = Decimal("0")
    netto_gefunden = False
    aktuell = ""
    im_fuss = "summe_pos" not in p

    def passt(feld, zeile):
        return re.search(p[feld], zeile) if feld in p else None

    for zeile in text.splitlines():
        s = zeile.strip()
        # Unterstreichungslinien überlagern im PDF manchmal den Text -> für Summenzeilen entfernen
        u = re.sub(r"\s+", " ", s.replace("_", "")).strip()
        if passt("reset", s):
            aktuell = ""
        elif m := passt("kommission", s):
            aktuell = m.group(1).strip()
        elif m := passt("betrag", u):
            key = kommission_key(aktuell)
            k = kommissionen.setdefault(key, {"namen": [], "summe": Decimal("0")})
            k["summe"] += zahl(m.group(1))
            name = beschreibung(aktuell)
            if name and name.upper() not in [n.upper() for n in k["namen"]]:
                k["namen"].append(name)
        elif m := passt("summe_pos", u):
            summe_pos = zahl(m.group(1))
            im_fuss = True
        elif m := passt("netto", u):
            netto += zahl(m.group(1))
            netto_gefunden = True
            if passt("ende", u):
                break
        elif passt("ende", u):
            break
        elif im_fuss and not netto_gefunden and (m := passt("zuschlag", u)):
            zuschlaege.append((m.group(1).strip(), zahl(m.group(2))))

    return info, kommissionen, zuschlaege, summe_pos, (netto if netto_gefunden else None)


def anteilig(kommissionen, zuschlag_gesamt):
    """Verteilt den Zuschlag anteilig nach Zwischensumme, centgenau (Restcents nach größtem Rest)."""
    basis = sum(k["summe"] for k in kommissionen.values())
    cents = int(zuschlag_gesamt * 100)
    if basis == 0 or cents == 0:
        return {key: Decimal("0") for key in kommissionen}
    roh = {key: Decimal(cents) * k["summe"] / basis for key, k in kommissionen.items()}
    ganz = {key: int(v.to_integral_value(ROUND_DOWN)) for key, v in roh.items()}
    rest = cents - sum(ganz.values())
    for key in sorted(roh, key=lambda x: roh[x] - ganz[x], reverse=True)[:rest]:
        ganz[key] += 1
    return {key: Decimal(c) / 100 for key, c in ganz.items()}


def schreibe_excel(ziel, info, kommissionen, zuschlaege, summe_pos, netto):
    zuschlag_gesamt = sum((b for _, b in zuschlaege), Decimal("0"))
    anteile = anteilig(kommissionen, zuschlag_gesamt)
    summe_zw = sum((k["summe"] for k in kommissionen.values()), Decimal("0"))
    summe_neu = summe_zw + sum(anteile.values(), Decimal("0"))

    wb = Workbook()
    ws = wb.active
    ws.title = "Kommissionen"
    fett = Font(bold=True)
    kopf_fill = PatternFill("solid", fgColor="D9E1F2")
    linie = Border(top=Side(style="thin"))
    eur = '#,##0.00 "€"'

    ws["A1"] = f"{info['lieferant']} – Rechnung {info['nummer']} vom {info['datum']}"
    ws["A1"].font = Font(bold=True, size=13)
    zuschlag_text = ", ".join(f"{n} {str(b).replace('.', ',')} €" for n, b in zuschlaege) or "keine"
    ws["A2"] = f"Zuschläge (anteilig verteilt): {zuschlag_text}"

    kopf = ["Kommission", "Baustelle", "Zwischensummen", "Anteil Zuschläge", "Netto gesamt"]
    ws.append([])
    ws.append(kopf)
    for c in ws[4]:
        c.font, c.fill = fett, kopf_fill

    def sortkey(k):
        return (k == OHNE, k)

    for key in sorted(kommissionen, key=sortkey):
        k = kommissionen[key]
        ws.append([key, " / ".join(k["namen"]), float(k["summe"]), float(anteile[key]),
                   float(k["summe"] + anteile[key])])

    ws.append(["Summe", "", float(summe_zw), float(zuschlag_gesamt), float(summe_neu)])
    for c in ws[ws.max_row]:
        c.font, c.border = fett, linie

    # Kontrolle
    ok_pos = summe_pos is None or summe_zw == summe_pos
    ok_netto = netto is not None and summe_neu == netto
    ws.append([])
    ws.append(["Kontrolle"])
    ws[ws.max_row][0].font = fett
    if summe_pos is not None:
        ws.append(["Summe Positionen laut PDF", "", float(summe_pos), "", "OK" if ok_pos else "FEHLER"])
    ws.append(["Nettosumme laut PDF", "", "", "", float(netto) if netto is not None else "nicht gefunden"])
    ws.append(["Abgleich Netto gesamt", "", "", "", "OK" if ok_netto else "FEHLER"])

    for zeile in ws.iter_rows(min_row=5):
        for c in zeile[2:5]:
            if isinstance(c.value, float):
                c.number_format = eur
    for spalte, breite in zip("ABCDE", (16, 45, 17, 17, 17)):
        ws.column_dimensions[spalte].width = breite
    ws.freeze_panes = "A5"

    wb.save(ziel)
    return ok_pos and ok_netto, summe_zw, summe_neu, summe_pos, netto


def verarbeite(pfad):
    text = lies_text(pfad)
    if not text.strip():
        raise ValueError("Kein Text im PDF (gescannt?) - Texterkennung ist noch nicht eingebaut.")
    profil = finde_profil(text)
    if profil is None:
        raise ValueError("Unbekannter Lieferant - dafür gibt es noch kein Profil.")
    info, kommissionen, zuschlaege, summe_pos, netto = auswerten(text, profil)
    if not kommissionen:
        raise ValueError(f"Keine Beträge gefunden (Profil {profil['name']}).")
    ziel = os.path.splitext(pfad)[0] + "_Kommissionen.xlsx"
    ok, *_ = schreibe_excel(ziel, info, kommissionen, zuschlaege, summe_pos, netto)
    return ziel, ok


def main():
    dateien = [a for a in sys.argv[1:] if a.lower().endswith(".pdf")]
    gui = not os.environ.get("RA_NOGUI")
    if gui:
        import tkinter as tk
        from tkinter import filedialog, messagebox
        root = tk.Tk()
        root.withdraw()
        if not dateien:
            dateien = list(filedialog.askopenfilenames(title="Rechnung(en) auswählen",
                                                       filetypes=[("PDF", "*.pdf")]))
        if not dateien:
            return

    meldungen, fehler = [], False
    for pfad in dateien:
        name = os.path.basename(pfad)
        try:
            ziel, ok = verarbeite(pfad)
            if ok:
                meldungen.append(f"✔ {name}: OK, Summen stimmen.")
            else:
                fehler = True
                meldungen.append(f"⚠ {name}: Summen stimmen NICHT mit dem PDF überein - bitte prüfen!")
            if gui and len(dateien) == 1:
                os.startfile(ziel)
        except Exception as e:  # noqa: BLE001
            fehler = True
            meldungen.append(f"✖ {name}: {e}")

    text = "\n".join(meldungen)
    if gui:
        (messagebox.showwarning if fehler else messagebox.showinfo)("Rechnungsauswertung", text)
    else:
        print(text)


if __name__ == "__main__":
    main()
