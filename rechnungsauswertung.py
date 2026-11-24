"""Rechnungsauswertung - fasst die Beträge von Lieferantenrechnungen pro Kommission zusammen.

Lieferanten: siehe PROFILE weiter unten.

Bedienung:
  - PDF(s) auf die .exe ziehen, oder
  - .exe doppelklicken und PDF(s) auswählen.
Es öffnet sich eine temporäre Excel-Datei (wird nirgends neben den PDFs gespeichert):
ein Blatt pro Rechnung, bei mehreren Rechnungen zusätzlich das Blatt "Baustellen".
"""
import os
import re
import sys
import tempfile
from collections import OrderedDict
from datetime import datetime
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
    text = (text or "").strip()
    # 5-stellige Nummer am Anfang; sonst die erste alleinstehende 5-stellige Nummer im Text
    m = re.match(r"(\d{5})(?!\d)", text) or re.search(r"(?<!\d)(\d{5})(?!\d)", text)
    if m:
        return m.group(1)
    # keine Kommissionsnummer, aber ein Text (z. B. heroal-Bestellnummer "122257 ART BTV"):
    # als eigene Zeile führen - Schlüssel ist die führende Nummer, sonst der ganze Text
    m = re.match(r"(\d+)\b", text)
    return m.group(1) if m else (text or OHNE)


def beschreibung(text):
    # "26010_RK SCHAFFERER WA LANS" / "26010/MKO X" -> "SCHAFFERER WA LANS" / "X"
    text = (text or "").strip()
    if re.match(r"\d{5}(?!\d)", text):
        return re.sub(r"^\d{5}(?:[_/]?[A-Za-z]{2,4}(?=\s|$))?\s*[-–]?\s*", "", text).strip()
    if re.search(r"(?<!\d)\d{5}(?!\d)", text):
        return text
    m = re.match(r"\d+\b\s*(.*)", text)
    return m.group(1) if m else ""


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


FETT = Font(bold=True)
KOPF_FILL = PatternFill("solid", fgColor="D9E1F2")
LINIE = Border(top=Side(style="thin"))
EUR = '#,##0.00 "€"'


def sortkey(k):
    return (k == OHNE, k)


def fuelle_blatt(ws, info, kommissionen, zuschlaege, summe_pos, netto):
    """Schreibt die Auswertung einer Rechnung in ein Blatt. Rückgabe: (ok, {Kommission: Netto gesamt})."""
    zuschlag_gesamt = sum((b for _, b in zuschlaege), Decimal("0"))
    anteile = anteilig(kommissionen, zuschlag_gesamt)
    summe_zw = sum((k["summe"] for k in kommissionen.values()), Decimal("0"))
    summe_neu = summe_zw + sum(anteile.values(), Decimal("0"))

    ws["A1"] = f"{info['lieferant']} – Rechnung {info['nummer']} vom {info['datum']}"
    ws["A1"].font = Font(bold=True, size=13)
    zuschlag_text = ", ".join(f"{n} {str(b).replace('.', ',')} €" for n, b in zuschlaege) or "keine"
    ws["A2"] = f"Zuschläge (anteilig verteilt): {zuschlag_text}"

    ws.append([])
    ws.append(["Kommission", "Baustelle", "Zwischensummen", "Anteil Zuschläge", "Netto gesamt"])
    for c in ws[4]:
        c.font, c.fill = FETT, KOPF_FILL

    for key in sorted(kommissionen, key=sortkey):
        k = kommissionen[key]
        ws.append([key, " / ".join(k["namen"]), float(k["summe"]), float(anteile[key]),
                   float(k["summe"] + anteile[key])])

    ws.append(["Summe", "", float(summe_zw), float(zuschlag_gesamt), float(summe_neu)])
    for c in ws[ws.max_row]:
        c.font, c.border = FETT, LINIE

    # Kontrolle
    ok_pos = summe_pos is None or summe_zw == summe_pos
    ok_netto = netto is not None and summe_neu == netto
    ws.append([])
    ws.append(["Kontrolle"])
    ws[ws.max_row][0].font = FETT
    if summe_pos is not None:
        ws.append(["Summe Positionen laut PDF", "", float(summe_pos), "", "OK" if ok_pos else "FEHLER"])
    ws.append(["Nettosumme laut PDF", "", "", "", float(netto) if netto is not None else "nicht gefunden"])
    ws.append(["Abgleich Netto gesamt", "", "", "", "OK" if ok_netto else "FEHLER"])

    for zeile in ws.iter_rows(min_row=5):
        for c in zeile[2:5]:
            if isinstance(c.value, float):
                c.number_format = EUR
    for spalte, breite in zip("ABCDE", (16, 45, 17, 17, 17)):
        ws.column_dimensions[spalte].width = breite
    ws.freeze_panes = "A5"

    return ok_pos and ok_netto, {key: (kommissionen[key]["namen"], kommissionen[key]["summe"] + anteile[key])
                                 for key in kommissionen}


def fuelle_baustellen(ws, rechnungen, ausgelassen):
    """Sammelblatt: pro Kommission jede Rechnung mit Betrag, darunter die Summe der Kommission."""
    ws["A1"] = "Baustellen – alle Rechnungen zusammen"
    ws["A1"].font = Font(bold=True, size=13)
    ws["A2"] = ("Nicht enthalten (Summen stimmen nicht): " + ", ".join(ausgelassen)) if ausgelassen \
        else "Alle gelesenen Rechnungen enthalten."
    if ausgelassen:
        ws["A2"].font = Font(bold=True, color="C00000")
    ws.append([])
    ws.append(["Kommission", "Baustelle", "Lieferant", "Rechnung", "Datum", "Netto"])
    for c in ws[4]:
        c.font, c.fill = FETT, KOPF_FILL

    je_kom = {}
    for info, betraege in rechnungen:
        for key, (namen, betrag) in betraege.items():
            je_kom.setdefault(key, []).append((info, namen, betrag))

    gesamt = Decimal("0")
    for key in sorted(je_kom, key=sortkey):
        summe = Decimal("0")
        for info, namen, betrag in je_kom[key]:
            ws.append([key, " / ".join(namen), info["lieferant"], info["nummer"], info["datum"], float(betrag)])
            summe += betrag
        ws.append([f"Summe {key}", "", "", "", "", float(summe)])
        for c in ws[ws.max_row]:
            c.font, c.border = FETT, LINIE
        ws.append([])
        gesamt += summe
    ws.append(["Gesamt", "", "", "", "", float(gesamt)])
    for c in ws[ws.max_row]:
        c.font, c.border = FETT, LINIE
    ws.append(["Summe der Rechnungen (Kontrolle)", "", "", "", "",
               float(sum((sum((b for _, b in bt.values()), Decimal("0")) for _, bt in rechnungen), Decimal("0")))])

    for zeile in ws.iter_rows(min_row=5):
        if isinstance(zeile[5].value, float):
            zeile[5].number_format = EUR
    for spalte, breite in zip("ABCDEF", (30, 45, 16, 16, 12, 17)):
        ws.column_dimensions[spalte].width = breite
    ws.freeze_panes = "A5"


def lies_rechnung(pfad):
    text = lies_text(pfad)
    if not text.strip():
        raise ValueError("Kein Text im PDF (gescannt?) - Texterkennung ist noch nicht eingebaut.")
    profil = finde_profil(text)
    if profil is None:
        raise ValueError("Unbekannter Lieferant - dafür gibt es noch kein Profil.")
    daten = auswerten(text, profil)
    if not daten[1]:
        raise ValueError(f"Keine Beträge gefunden (Profil {profil['name']}).")
    return daten


def blattname(wb, info):
    name = re.sub(r"[\[\]:*?/\\]", "-", f"{info['lieferant'][:12]} {info['nummer']}".strip())[:31] or "Rechnung"
    basis, n = name, 2
    while name in wb.sheetnames:
        name = f"{basis[:28]}_{n}"
        n += 1
    return name


def erstelle_mappe(dateien):
    """Liest alle PDFs. Rückgabe: (Workbook oder None, Meldungen, Fehler ja/nein)."""
    wb = Workbook()
    wb.remove(wb.active)
    meldungen, fehler = [], False
    rechnungen, ausgelassen, gesehen = [], [], set()
    for pfad in dateien:
        name = os.path.basename(pfad)
        try:
            info, kommissionen, zuschlaege, summe_pos, netto = lies_rechnung(pfad)
            kennung = (info["lieferant"], info["nummer"])
            if info["nummer"] and kennung in gesehen:
                meldungen.append(f"⚠ {name}: Rechnung {info['nummer']} ist doppelt - nur einmal gezählt.")
                fehler = True
                continue
            gesehen.add(kennung)
            ws = wb.create_sheet(blattname(wb, info))
            ok, betraege = fuelle_blatt(ws, info, kommissionen, zuschlaege, summe_pos, netto)
            if ok:
                rechnungen.append((info, betraege))
                meldungen.append(f"✔ {name}: OK, Summen stimmen.")
            else:
                fehler = True
                ausgelassen.append(f"{info['lieferant']} {info['nummer']}")
                meldungen.append(f"⚠ {name}: Summen stimmen NICHT mit dem PDF überein - bitte prüfen!")
        except Exception as e:  # noqa: BLE001
            fehler = True
            meldungen.append(f"✖ {name}: {e}")
    if not wb.sheetnames:
        return None, meldungen, fehler
    if len(wb.sheetnames) > 1:
        fuelle_baustellen(wb.create_sheet("Baustellen"), rechnungen, ausgelassen)
    return wb, meldungen, fehler


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

    wb, meldungen, fehler = erstelle_mappe(dateien)
    text = "\n".join(meldungen)
    ziel = None
    if wb is not None:
        # nur temporär: Excel öffnet die Datei, gespeichert wird nur, wenn man es in Excel selbst tut
        ziel = os.path.join(tempfile.mkdtemp(prefix="Rechnungsauswertung_"),
                            f"Rechnungsauswertung_{datetime.now():%Y-%m-%d_%H%M%S}.xlsx")
        wb.save(ziel)
    if gui:
        if fehler or ziel is None:
            messagebox.showwarning("Rechnungsauswertung", text)
        if ziel:
            os.startfile(ziel)
    else:
        print(text)
        print(ziel)


if __name__ == "__main__":
    main()
