# Rechnungsauswertung

Fasst die Zwischensummen von Lieferantenrechnungen (digitale PDFs; derzeit Schachermayer, heroal, Vetrotech, Süd-Metall) pro Kommission (5-stellige Nummer) zusammen, verteilt Zuschläge anteilig und prüft gegen Summe Positionen und Nettosumme.

**Bedienung:** PDF(s) auf `Rechnungsauswertung.exe` ziehen oder EXE starten und PDFs auswählen. Es öffnet sich eine temporäre Excel-Datei (nichts wird neben den PDFs gespeichert): ein Blatt pro Rechnung, bei mehreren Rechnungen zusätzlich das Blatt *Baustellen* mit allen Beträgen je Kommission samt Rechnungsnummer und Summe.

Die EXE baut GitHub Actions bei jedem Push (Reiter *Actions* → letzter Lauf → *Artifacts*).
