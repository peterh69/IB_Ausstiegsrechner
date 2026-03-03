# IB Ausstiegsrechner

Liest alle offenen Positionen aus der Interactive Brokers Trader Workstation (TWS) aus
und erstellt eine übersichtliche, farblich strukturierte Excel-Analyse.

## Features

- **Guthaben-Übersicht**: EUR/USD Gesamtguthaben, für CSPs gebundenes Kapital, freier Cash
- **Cash Secured Puts (CSPs)**: Strike, DTE, Prämie, annualisierte Restrendite
- **Aktienbestände**: Aktueller Kurs, Einstandspreis – getrennt nach EUR und USD
- **Covered Calls**: Strike, DTE, Prämie, annualisierte Restrendite
- **Genaue Bezeichnung**: Vollständiger Firmenname für alle Positionen (z.B. „NVIDIA Corporation")
- **EUR/USD-Kurs** im Datei-Header
- Farblich strukturierte Excel-Ausgabe mit automatischer Spaltenbreite

## Restrendite-Berechnung

Die annualisierte Restrendite wird nur angezeigt, wenn eine **Gewinnposition** vorliegt
(aktueller Optionspreis < erhaltene Prämie):

```
Restrendite p.a. = (365 / DTE) × (Erhaltene Prämie / Strike)
```

## Voraussetzungen

- Interactive Brokers TWS oder IB Gateway muss laufen
- API muss in TWS aktiviert sein:
  `File → Global Configuration → API → Settings → Enable ActiveX and Socket Clients`
- Python 3.10+

## Installation

```bash
pip install -r requirements.txt
```

## Konfiguration

In `ausstiegsrechner.py` am Anfang der Datei:

| Parameter          | Standard            | Beschreibung                                       |
|--------------------|---------------------|----------------------------------------------------|
| `TWS_HOST`         | `127.0.0.1`         | Hostname/IP der TWS-Instanz                        |
| `TWS_PORT`         | `7496`              | API-Port (7496=Live, 7497=Paper, 4001=Gateway)     |
| `CLIENT_ID`        | `10`                | Eindeutige Client-ID für diese Verbindung          |
| `OUTPUT_FILE`      | `IB_Positionen.xlsx`| Name der Excel-Ausgabedatei                        |
| `MARKET_DATA_WAIT` | `3`                 | Wartezeit in Sekunden für Marktdaten               |

## Verwendung

```bash
python ausstiegsrechner.py
```

## Freier Cash

```
Freier Cash = Gesamtguthaben (je Währung) − für CSPs gebundenes Kapital
Gebundenes Kapital je CSP = Strike × 100 × |Anzahl Kontrakte|
```

## Ausgabe-Struktur

### Abschnitt 0: Guthaben-Übersicht

|                | EUR    | USD    |
|----------------|--------|--------|
| Gesamtguthaben | ...    | ...    |
| CSP-Kapital    | ...    | ...    |
| Freier Cash    | ...    | ...    |

### Abschnitt 1: Cash Secured Puts

| Symbol | Bezeichnung | Position | Kurs Underlying | Strike | DTE | Ablaufdatum | Erhaltene Prämie | Akt. Options-Preis | Restrendite p.a. | Währung |

### Abschnitt 2: Aktien & Covered Calls

Getrennt in **EUR-Positionen** und **USD-Positionen**. Je Symbol: Aktienzeile, dann zugehörige Calls.

| Symbol | Bezeichnung | Typ | Position | Akt. Kurs | Strike | DTE | Ablaufdatum | Kauf-/Verkaufspreis | Akt. Options-Preis | Restrendite p.a. | Währung |
