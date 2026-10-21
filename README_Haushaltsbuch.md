# Haushaltsbuch mit PySide6

## Installation

```bash
python -m venv .venv
# Windows:
.venv\Scripts\activate
# macOS/Linux:
source .venv/bin/activate
pip install -r requirements.txt
python haushaltsbuch.py
```

## Bedienhinweise

- Positive Beträge sind Einnahmen, negative Beträge sind Ausgaben.
- Mit `Tab` wechseln Sie durch die Eingabefelder.
- Nach dem Betrag fügt `Enter` die Buchung hinzu. Kategorie und Konto bleiben ausgewählt.
- Ein Doppelklick auf eine Tabellenzeile öffnet die Bearbeitung.
- JSON speichert Buchungen und Einstellungen gemeinsam.
- CSV/TXT-Import erwartet die Spalten: `Datum`, `Kategorie`, `Beschreibung`, `PayPal`, `Regelmäßig`, `Konto`, `Betrag`.
- Das Datum wird als `YYYY-MM-DD` importiert/exportiert.
- Buchungen mit Beschreibung `Umbuchung` werden in den Ein-/Ausgangsdiagrammen ausgeschlossen.

## Datenmodell

Die Konten besitzen einen Anfangsstand. Der Kontostand ergibt sich aus Anfangsstand plus Buchungen. Für einen produktiven Mehrbenutzerbetrieb wäre die Nutzung über eine SQLite-Datenbank statt JSON sinnvoll.
