# Codex-Test
Erstes Repository für ChatGPT-Codex Umgebung ✅

## Vergleich von Nachunternehmer-Angeboten

Mit `scripts/compare_offers.py` lässt sich eine Excel-Auswertung aus mehreren
Angeboten erzeugen. Das Skript liest CSV- oder Excel-Dateien ein, gleicht die
Positionen anhand von LV-Nummer oder Bezeichnung ab und markiert pro Position
den günstigsten Anbieter inklusive Preisabweichungen.

```bash
python3 scripts/compare_offers.py ergebnis.xlsx angebot1.xlsx angebot2.csv angebot3.xlsx
```

Optional können mit `--keys` die Spaltennamen für die Zuordnung vorgegeben
werden, falls die automatische Erkennung nicht greift:

```bash
python3 scripts/compare_offers.py ergebnis.xlsx angebot1.xlsx angebot2.csv --keys "LV-Nr." Bezeichnung
```

Die benötigten Python-Pakete sind in `requirements.txt` hinterlegt.
