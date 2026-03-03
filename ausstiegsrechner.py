"""
IB Ausstiegsrechner - Interactive Brokers Portfolio Analyzer
============================================================

Dieses Programm verbindet sich mit der Interactive Brokers Trader Workstation (TWS)
über die offizielle IB API und liest alle offenen Positionen aus. Die Ergebnisse
werden aufbereitet und in einer übersichtlichen Excel-Datei gespeichert.

Unterstützte Positionstypen
---------------------------
- Cash Secured Puts (CSPs): Short-Put-Optionen mit Strike, DTE, Prämie und Restrendite
- Aktien (STK): Aktienpositionen mit aktuellem Kurs und Einstandspreis
- Covered Calls (CC): Short-Call-Optionen mit Strike, DTE, Prämie und Restrendite

Sortierung der Excel-Ausgabe
-----------------------------
1. Abschnitt: Alle CSPs, sortiert nach Symbol (alphabetisch) und DTE (aufsteigend)
2. Abschnitt: Aktien mit zugehörigen Covered Calls, alphabetisch nach Symbol

Restrendite-Berechnung
----------------------
Die annualisierte Restrendite wird nur berechnet, wenn eine Gewinnposition vorliegt
(aktueller Optionspreis < erhaltene Prämie):

    Restrendite p.a. = (365 / DTE) × (Erhaltene Prämie / Strike)

Voraussetzungen
---------------
1. Interactive Brokers TWS oder IB Gateway muss gestartet sein.
2. Die API-Verbindung muss in TWS aktiviert sein:
   File → Global Configuration → API → Settings → „Enable ActiveX and Socket Clients"
3. Python-Pakete: ib_insync, openpyxl (siehe requirements.txt)

Installation
------------
    pip install -r requirements.txt

Verwendung
----------
    python ausstiegsrechner.py

Konfiguration
-------------
Die Verbindungsparameter (Host, Port, Client-ID) können am Anfang der Datei
unter „Konfiguration" angepasst werden.

Ports:
    7496 = TWS Live-Trading
    7497 = TWS Paper-Trading
    4001 = IB Gateway Live
    4002 = IB Gateway Paper
"""

import asyncio
import logging
import sys
import time
from datetime import datetime, date

# Python 3.10+ erstellt keine Event Loop mehr automatisch.
# Vor dem Import von ib_insync muss eine neue Loop angelegt werden,
# damit asyncio-basierte Funktionen von ib_insync korrekt funktionieren.
asyncio.set_event_loop(asyncio.new_event_loop())

from ib_insync import IB, Forex, Stock, Option

from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ---------------------------------------------------------------------------
# Log-Filter: Harmlose IB-Fehlercodes unterdrücken
# ---------------------------------------------------------------------------

# Diese Fehlercodes sind von IB als reine Informationsmeldungen gedacht
# und stellen keine echten Fehler dar. Sie werden aus dem Log herausgefiltert,
# um die Konsolenausgabe übersichtlich zu halten.
_BENIGN_LOG_CODES = {321, 354, 2104, 2106, 2107, 2108, 2158, 10090, 10167}


class _IBLogFilter(logging.Filter):
    """Unterdrückt bekannte, harmlose ib_insync-Fehlermeldungen im Log."""

    def filter(self, record):
        """Gibt False zurück wenn die Meldung einen harmlosen Fehlercode enthält."""
        msg = record.getMessage()
        return not any(
            f'Error {c},' in msg or f'{c},' in msg[:20]
            for c in _BENIGN_LOG_CODES
        )


# Filter auf alle relevanten ib_insync-Logger anwenden
for _log_name in ('ib_insync.wrapper', 'ib_insync.client', 'ib_insync.ib'):
    logging.getLogger(_log_name).addFilter(_IBLogFilter())


# ---------------------------------------------------------------------------
# Konfiguration
# ---------------------------------------------------------------------------

TWS_HOST = '127.0.0.1'    # Hostname oder IP-Adresse der TWS/Gateway-Instanz
TWS_PORT = 7496            # API-Port (7496=TWS Live, 7497=TWS Paper, 4001=Gateway Live)
CLIENT_ID = 10             # Eindeutige Client-ID; darf nicht doppelt vergeben sein
OUTPUT_FILE = 'IB_Positionen.xlsx'  # Dateiname der Excel-Ausgabe
MARKET_DATA_WAIT = 3       # Sekunden, die auf eingehende Marktdaten gewartet wird


# ---------------------------------------------------------------------------
# Farbpalette für die Excel-Formatierung
# ---------------------------------------------------------------------------

COLOR_RED_HEADER   = 'C0504D'   # Dunkelrot  – Überschrift CSP-Abschnitt
COLOR_BLUE_HEADER  = '17375E'   # Dunkelblau – Überschrift Aktien/Calls-Abschnitt
COLOR_BLUE_STOCK   = 'DCE6F1'   # Hellblau   – Hintergrund Aktienzeile
COLOR_GREEN_CALL   = 'EBF1DE'   # Hellgrün   – Hintergrund Call-Zeile
COLOR_SEPARATOR    = 'D9D9D9'   # Grau       – Trennzeile zwischen Symbolen
COLOR_CSP_ROW_1    = 'FFFFFF'   # Weiß       – CSP-Zeile (gerader Index)
COLOR_CSP_ROW_2    = 'F2DCDB'   # Rosa       – CSP-Zeile (ungerader Index, alternierend)


# ---------------------------------------------------------------------------
# Hilfsfunktionen – Preisermittlung und Datumsformatierung
# ---------------------------------------------------------------------------

def get_price(ticker) -> float | None:
    """Ermittelt den besten verfügbaren Marktpreis aus einem ib_insync-Ticker.

    Prioritätsreihenfolge: last → close → Mitte aus Bid/Ask.
    Gibt None zurück, wenn kein Preis verfügbar ist.

    Args:
        ticker: ib_insync Ticker-Objekt oder None.

    Returns:
        Preis als float, oder None wenn kein Preis verfügbar.
    """
    if ticker is None:
        return None
    last  = ticker.last
    close = ticker.close
    bid   = ticker.bid
    ask   = ticker.ask

    if last and last > 0:
        return last
    if close and close > 0:
        return close
    if bid and ask and bid > 0 and ask > 0:
        return (bid + ask) / 2
    return None


def dte(expiry_str: str) -> int:
    """Berechnet die verbleibenden Tage bis zum Verfall (Days to Expiration).

    Args:
        expiry_str: Verfallsdatum im Format YYYYMMDD.

    Returns:
        Anzahl der verbleibenden Tage als int, oder 0 bei ungültigem Datum.
    """
    try:
        exp_date = datetime.strptime(expiry_str, '%Y%m%d').date()
        return (exp_date - date.today()).days
    except Exception:
        return 0


def fmt_date(expiry_str: str) -> str:
    """Konvertiert ein Datum von YYYYMMDD nach DD.MM.YYYY (deutsches Format).

    Args:
        expiry_str: Datum als String im Format YYYYMMDD.

    Returns:
        Datum als String im Format DD.MM.YYYY, oder den Originalwert bei Fehler.
    """
    try:
        return datetime.strptime(expiry_str, '%Y%m%d').strftime('%d.%m.%Y')
    except Exception:
        return expiry_str


# ---------------------------------------------------------------------------
# Hilfsfunktionen – Excel-Formatierung
# ---------------------------------------------------------------------------

def apply_fill(cell, hex_color: str):
    """Setzt die Hintergrundfarbe einer Excel-Zelle.

    Args:
        cell: openpyxl-Cell-Objekt.
        hex_color: Farbe als 6-stelliger Hex-String (ohne führendes #).
    """
    cell.fill = PatternFill(fill_type='solid', fgColor=hex_color)


def apply_header_style(cell, hex_color: str):
    """Formatiert eine Zelle als Abschnitts-Header (farbiger Hintergrund, weiße Fettschrift).

    Args:
        cell: openpyxl-Cell-Objekt.
        hex_color: Hintergrundfarbe als 6-stelliger Hex-String.
    """
    cell.fill = PatternFill(fill_type='solid', fgColor=hex_color)
    cell.font = Font(bold=True, color='FFFFFF')
    cell.alignment = Alignment(horizontal='center', vertical='center')


def thin_border() -> Border:
    """Erstellt einen dünnen grauen Rahmen für Trennzeilen.

    Returns:
        openpyxl Border-Objekt mit dünner unterer Linie.
    """
    thin = Side(style='thin', color='AAAAAA')
    return Border(bottom=thin)


# ---------------------------------------------------------------------------
# Hilfsfunktionen – Finanzberechnungen
# ---------------------------------------------------------------------------

def calc_restrendite(premium: float, strike: float, days: int,
                     current_price) -> float | None:
    """Berechnet die annualisierte Restrendite einer Short-Option-Position.

    Die Restrendite wird nur berechnet, wenn:
    - Ein gültiger Strike, eine positive Prämie und verbleibende Laufzeit vorliegen.
    - Der aktuelle Optionspreis unter der erhaltenen Prämie liegt (Gewinnposition).

    Formel:
        Restrendite p.a. = (365 / DTE) × (Erhaltene Prämie / Strike)

    Args:
        premium: Erhaltene Prämie pro Aktie (Einstiegsverkaufspreis).
        strike: Strike-Preis der Option.
        days: Verbleibende Tage bis zum Verfall (DTE).
        current_price: Aktueller Marktpreis der Option (zum Rückkauf), oder None.

    Returns:
        Restrendite als Dezimalzahl (z.B. 0.15 = 15% p.a.), oder None wenn
        keine Gewinnposition vorliegt oder die Berechnung nicht möglich ist.
    """
    if days <= 0 or strike <= 0 or premium <= 0:
        return None
    # Kein Gewinn: aktueller Preis >= erhaltene Prämie (Rückkauf wäre teurer)
    if current_price is not None and current_price >= premium:
        return None
    return (365.0 / days) * (premium / strike)


def fetch_long_names(ib: IB, stock_contracts: dict) -> dict:
    """Lädt die offiziellen Langnamen (Firmennamen) für alle übergebenen Kontrakte.

    Verwendet IB's reqContractDetails() um den vollständigen Firmennamen
    (z.B. 'NVIDIA Corporation') aus den Vertragsdaten zu lesen.

    Args:
        ib: Aktive IB-Verbindung.
        stock_contracts: Dict {symbol: Stock-Kontrakt}.

    Returns:
        Dict {symbol: long_name}. Bei Fehlern wird das Ticker-Symbol als Fallback verwendet.
    """
    long_names = {}
    for sym, contract in stock_contracts.items():
        try:
            details = ib.reqContractDetails(contract)
            if details and details[0].longName:
                long_names[sym] = details[0].longName
            else:
                long_names[sym] = sym
        except Exception:
            long_names[sym] = sym
    return long_names


# ---------------------------------------------------------------------------
# Hauptprogramm
# ---------------------------------------------------------------------------

def main():
    """Einstiegspunkt: Verbindet mit TWS und startet die Positions-Analyse."""
    ib = IB()
    print(f"Verbinde mit TWS auf {TWS_HOST}:{TWS_PORT} (clientId={CLIENT_ID})...")

    # Fehlercodes, die von IB als harmlos/informativ eingestuft werden
    SUPPRESS_CODES = _BENIGN_LOG_CODES | {2104, 2106, 2107, 2108, 2158}

    def error_handler(req_id, code, msg, contract=None):
        """Unterdrückt bekannte harmlose TWS-Statusmeldungen in der Konsolenausgabe."""
        if code in SUPPRESS_CODES:
            return
        print(f"TWS Fehler {code} (reqId {req_id}): {msg}")

    ib.errorEvent += error_handler

    try:
        ib.connect(TWS_HOST, TWS_PORT, clientId=CLIENT_ID, timeout=10, readonly=True)
    except Exception as e:
        print(f"FEHLER: Verbindung zu TWS fehlgeschlagen: {e}")
        print("Bitte sicherstellen, dass TWS läuft und die API aktiviert ist.")
        sys.exit(1)

    print("Verbunden.")
    try:
        _run(ib)
    finally:
        ib.disconnect()
        print("Verbindung getrennt.")


def _run(ib: IB):
    """Liest alle Positionen aus und erstellt die Excel-Datei.

    Diese Funktion orchestriert den gesamten Prozess:
    1. EUR/USD-Wechselkurs abrufen
    2. Positionen laden und klassifizieren (CSPs, Aktien, Calls)
    3. Kontrakte qualifizieren und Marktdaten gebündelt abrufen
    4. Firmennamen (Genaue Bezeichnung) für alle Symbole laden
    5. Excel-Datei mit zwei Abschnitten aufbauen und speichern

    Args:
        ib: Aktive, bereits verbundene IB-Instanz.
    """
    # --- EUR/USD-Wechselkurs anfordern ---
    # Wird für die Anzeige im Datei-Header benötigt
    eurusd_contract = Forex('EURUSD')
    ib.qualifyContracts(eurusd_contract)
    eurusd_ticker = ib.reqMktData(eurusd_contract, '', False, False)

    # --- Positionen laden ---
    positions = ib.positions()
    if not positions:
        print("Keine offenen Positionen gefunden.")
        return

    # --- Positionen klassifizieren ---
    csps   = []   # Short Puts (Cash Secured Puts)
    stocks = []   # Aktienpositionen
    calls  = []   # Short Calls (Covered Calls)

    for pos in positions:
        contract = pos.contract
        sec_type = contract.secType
        right = getattr(contract, 'right', '')

        if sec_type == 'OPT' and right == 'P':
            csps.append(pos)
        elif sec_type == 'OPT' and right == 'C':
            calls.append(pos)
        elif sec_type == 'STK':
            stocks.append(pos)

    # --- Underlying-Kontrakte für Optionen erzeugen ---
    # Wird benötigt um den aktuellen Kurs des Basiswerts und den Firmennamen abzufragen
    underlying_symbols = set()
    for pos in csps + calls:
        c = pos.contract
        underlying_symbols.add((c.symbol, getattr(c, 'currency', 'USD')))

    underlying_contracts = {}
    for sym, cur in underlying_symbols:
        stk = Stock(sym, 'SMART', cur)
        try:
            ib.qualifyContracts(stk)
        except Exception:
            pass
        underlying_contracts[(sym, cur)] = stk

    # --- Marktdaten gebündelt anfordern ---

    # Optionen: qualifizieren und Marktdaten abonnieren
    opt_tickers = {}
    for pos in csps + calls:
        c = pos.contract
        try:
            ib.qualifyContracts(c)
        except Exception:
            pass
        t = ib.reqMktData(c, '', False, False)
        opt_tickers[c.conId] = t

    # Aktien: qualifizieren und Marktdaten abonnieren
    stk_tickers = {}
    for pos in stocks:
        c = pos.contract
        try:
            ib.qualifyContracts(c)
        except Exception:
            pass
        t = ib.reqMktData(c, '', False, False)
        stk_tickers[c.conId] = t

    # Underlying-Kurse für Optionen abonnieren
    underlying_tickers = {}
    for key, stk in underlying_contracts.items():
        t = ib.reqMktData(stk, '', False, False)
        underlying_tickers[key] = t

    # Auf Marktdaten warten
    print(f"Warte {MARKET_DATA_WAIT}s auf Marktdaten...")
    time.sleep(MARKET_DATA_WAIT)
    ib.sleep(0)  # ib_insync-Event-Loop einen Verarbeitungszyklus ausführen lassen

    # --- Firmennamen (Genaue Bezeichnung) laden ---
    # Alle bekannten Stock-Kontrakte zusammenführen: Underlyings von Optionen + direkte Aktien
    all_known_stocks = {}
    for (sym, _cur), contract in underlying_contracts.items():
        all_known_stocks[sym] = contract
    for pos in stocks:
        sym = pos.contract.symbol
        if sym not in all_known_stocks:
            all_known_stocks[sym] = pos.contract

    print("Lade Firmennamen...")
    long_names = fetch_long_names(ib, all_known_stocks)

    # --- EUR/USD-Kurs auslesen ---
    eurusd_price = get_price(eurusd_ticker)
    if eurusd_price is None:
        eurusd_price = 1.0
        print("WARNUNG: EUR/USD-Kurs nicht verfügbar, verwende 1.0")

    # ---------------------------------------------------------------------------
    # Daten aufbereiten
    # ---------------------------------------------------------------------------

    # CSP-Daten sammeln
    csp_rows = []
    for pos in csps:
        c = pos.contract
        sym = c.symbol
        cur = getattr(c, 'currency', 'USD')

        # avgCost bei Optionen = Prämie × Kontraktmultiplikator (100)
        # → Division durch 100 ergibt die Prämie pro Aktie/Anteil
        premium_per_share = abs(pos.avgCost) / 100.0

        opt_ticker      = opt_tickers.get(c.conId)
        current_opt_price = get_price(opt_ticker)

        ul_ticker       = underlying_tickers.get((sym, cur))
        underlying_price = get_price(ul_ticker)

        strike = float(getattr(c, 'strike', 0))
        expiry = getattr(c, 'lastTradeDateOrContractMonth', '')
        days   = dte(expiry)

        csp_rows.append({
            'symbol':           sym,
            'bezeichnung':      long_names.get(sym, sym),
            'position':         pos.position,   # negative Zahl bei Short-Position
            'underlying_price': underlying_price,
            'strike':           strike,
            'dte':              days,
            'expiry':           expiry,
            'premium':          premium_per_share,
            'current_price':    current_opt_price,
            'eurusd':           eurusd_price,
            'currency':         cur,
        })

    # Sortierung: Symbol alphabetisch, dann DTE aufsteigend
    csp_rows.sort(key=lambda r: (r['symbol'], r['dte']))

    # Aktien-Daten sammeln
    stock_map = {}
    for pos in stocks:
        c = pos.contract
        sym = c.symbol
        cur = getattr(c, 'currency', 'USD')

        stk_ticker    = stk_tickers.get(c.conId)
        current_price = get_price(stk_ticker)

        stock_map[sym] = {
            'symbol':        sym,
            'bezeichnung':   long_names.get(sym, sym),
            'position':      pos.position,
            'avg_cost':      pos.avgCost,   # Einstandspreis pro Aktie
            'current_price': current_price,
            'eurusd':        eurusd_price,
            'currency':      cur,
        }

    # Call-Daten sammeln
    call_map = {}   # {symbol: [call_row, ...]}
    for pos in calls:
        c = pos.contract
        sym = c.symbol
        cur = getattr(c, 'currency', 'USD')

        premium_per_share = abs(pos.avgCost) / 100.0

        opt_ticker      = opt_tickers.get(c.conId)
        current_opt_price = get_price(opt_ticker)

        strike = float(getattr(c, 'strike', 0))
        expiry = getattr(c, 'lastTradeDateOrContractMonth', '')
        days   = dte(expiry)

        row = {
            'symbol':        sym,
            'bezeichnung':   long_names.get(sym, sym),
            'position':      pos.position,
            'strike':        strike,
            'dte':           days,
            'expiry':        expiry,
            'premium':       premium_per_share,
            'current_price': current_opt_price,
            'eurusd':        eurusd_price,
            'currency':      cur,
        }
        call_map.setdefault(sym, []).append(row)

    # Calls innerhalb eines Symbols nach DTE aufsteigend sortieren
    for sym in call_map:
        call_map[sym].sort(key=lambda r: r['dte'])

    # Alle Symbole für Abschnitt 2 (Aktien + Calls), alphabetisch sortiert
    all_syms_2 = sorted(set(list(stock_map.keys()) + list(call_map.keys())))

    # ---------------------------------------------------------------------------
    # Excel-Datei aufbauen
    # ---------------------------------------------------------------------------

    wb = Workbook()
    ws = wb.active
    ws.title = 'Positionen'

    now_str = datetime.now().strftime('%d.%m.%Y %H:%M')
    current_row = 1

    # --- Titelzeile ---
    title_text = f"IB Positionen  |  Stand: {now_str}  |  EUR/USD: {eurusd_price:.4f}"
    title_cell = ws.cell(row=current_row, column=1, value=title_text)
    title_cell.font = Font(bold=True, size=13)
    ws.merge_cells(start_row=current_row, start_column=1,
                   end_row=current_row, end_column=12)
    current_row += 2

    # ===========================================================
    # ABSCHNITT 1: Cash Secured Puts (CSPs)
    # ===========================================================

    csp_headers = [
        'Symbol', 'Bezeichnung', 'Position', 'Kurs Underlying',
        'Strike', 'DTE', 'Ablaufdatum',
        'Erhaltene Prämie', 'Akt. Options-Preis', 'Restrendite p.a.', 'Währung'
    ]
    NUM_CSP_COLS = len(csp_headers)

    # Abschnitts-Header (rote Kopfzeile)
    ws.cell(row=current_row, column=1, value='Cash Secured Puts (CSPs)')
    for col in range(1, NUM_CSP_COLS + 1):
        apply_header_style(ws.cell(row=current_row, column=col), COLOR_RED_HEADER)
    ws.merge_cells(start_row=current_row, start_column=1,
                   end_row=current_row, end_column=NUM_CSP_COLS)
    current_row += 1

    # Spalten-Header
    for col_idx, header in enumerate(csp_headers, start=1):
        cell = ws.cell(row=current_row, column=col_idx, value=header)
        cell.font = Font(bold=True)
        cell.fill = PatternFill(fill_type='solid', fgColor='F2DCDB')
        cell.alignment = Alignment(horizontal='center')
    current_row += 1

    # Datenzeilen (alternierend gefärbt)
    for i, row in enumerate(csp_rows):
        bg = COLOR_CSP_ROW_1 if i % 2 == 0 else COLOR_CSP_ROW_2
        ul_price    = row['underlying_price']
        cur_opt     = row['current_price']
        restrendite = calc_restrendite(row['premium'], row['strike'],
                                       row['dte'], cur_opt)

        values = [
            row['symbol'],                                          # 1  Ticker Symbol
            row['bezeichnung'],                                     # 2  Genaue Bezeichnung
            row['position'],                                        # 3  Anzahl Kontrakte
            ul_price if ul_price is not None else 'n/v',           # 4  Kurs Underlying
            row['strike'],                                          # 5  Strike-Preis
            row['dte'],                                             # 6  DTE (Restlaufzeit in Tagen)
            fmt_date(row['expiry']),                                # 7  Ablaufdatum
            row['premium'],                                         # 8  Erhaltene Prämie pro Aktie
            cur_opt if cur_opt is not None else 'n/v',             # 9  Aktueller Options-Preis
            restrendite if restrendite is not None else '-',        # 10 Restrendite p.a.
            row['currency'],                                        # 11 Währung (USD/EUR)
        ]
        for col_idx, val in enumerate(values, start=1):
            cell = ws.cell(row=current_row, column=col_idx, value=val)
            apply_fill(cell, bg)
            cell.alignment = Alignment(horizontal='right' if col_idx > 2 else 'left')
            if col_idx in (4, 5, 8, 9) and isinstance(val, float):
                cell.number_format = '#,##0.00'
            elif col_idx == 10 and isinstance(val, float):
                cell.number_format = '0.00%'
        current_row += 1

    current_row += 1  # Leerzeile zwischen den Abschnitten

    # ===========================================================
    # ABSCHNITT 2: Aktien & Covered Calls
    # ===========================================================

    sec2_headers = [
        'Symbol', 'Bezeichnung', 'Typ', 'Position', 'Akt. Kurs',
        'Strike', 'DTE', 'Ablaufdatum',
        'Kauf-/Verkaufspreis', 'Akt. Options-Preis', 'Restrendite p.a.', 'Währung'
    ]
    NUM_SEC2_COLS = len(sec2_headers)

    # Abschnitts-Header (blauer Kopfzeile)
    ws.cell(row=current_row, column=1, value='Aktien & Covered Calls')
    for col in range(1, NUM_SEC2_COLS + 1):
        apply_header_style(ws.cell(row=current_row, column=col), COLOR_BLUE_HEADER)
    ws.merge_cells(start_row=current_row, start_column=1,
                   end_row=current_row, end_column=NUM_SEC2_COLS)
    current_row += 1

    # Spalten-Header
    for col_idx, header in enumerate(sec2_headers, start=1):
        cell = ws.cell(row=current_row, column=col_idx, value=header)
        cell.font = Font(bold=True)
        cell.fill = PatternFill(fill_type='solid', fgColor='DCE6F1')
        cell.alignment = Alignment(horizontal='center')
    current_row += 1

    # Datenzeilen: je Symbol erst Aktienzeile, dann zugehörige Call-Zeilen
    for sym_idx, sym in enumerate(all_syms_2):

        # --- Aktienzeile ---
        if sym in stock_map:
            s = stock_map[sym]
            cur_price = s['current_price']
            values = [
                sym,                                                # 1  Ticker Symbol
                s['bezeichnung'],                                   # 2  Genaue Bezeichnung
                'Aktie',                                            # 3  Typ
                s['position'],                                      # 4  Anzahl Aktien
                cur_price if cur_price is not None else 'n/v',     # 5  Aktueller Kurs
                '-', '-', '-',                                      # 6-8 Strike, DTE, Ablauf (n/a)
                s['avg_cost'],                                      # 9  Einstandspreis pro Aktie
                '-', '-',                                           # 10-11 Opt-Preis, Restrendite (n/a)
                s['currency'],                                      # 12 Währung
            ]
            for col_idx, val in enumerate(values, start=1):
                cell = ws.cell(row=current_row, column=col_idx, value=val)
                apply_fill(cell, COLOR_BLUE_STOCK)
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal='right' if col_idx > 2 else 'left')
                if col_idx == 5 and isinstance(val, float):
                    cell.number_format = '#,##0.00'
                if col_idx == 9 and isinstance(val, float):
                    cell.number_format = '#,##0.00'
            current_row += 1

        # --- Call-Zeilen (nach DTE aufsteigend sortiert) ---
        if sym in call_map:
            for call_row in call_map[sym]:
                cur_opt     = call_row['current_price']
                restrendite = calc_restrendite(
                    call_row['premium'], call_row['strike'],
                    call_row['dte'], cur_opt
                )
                values = [
                    sym,                                            # 1  Ticker Symbol
                    call_row['bezeichnung'],                        # 2  Genaue Bezeichnung
                    'Call',                                         # 3  Typ
                    call_row['position'],                           # 4  Anzahl Kontrakte
                    '-',                                            # 5  Akt. Kurs (n/a bei Call)
                    call_row['strike'],                             # 6  Strike-Preis
                    call_row['dte'],                                # 7  DTE
                    fmt_date(call_row['expiry']),                   # 8  Ablaufdatum
                    call_row['premium'],                            # 9  Erhaltene Prämie
                    cur_opt if cur_opt is not None else 'n/v',     # 10 Aktueller Options-Preis
                    restrendite if restrendite is not None else '-',# 11 Restrendite p.a.
                    call_row['currency'],                           # 12 Währung
                ]
                for col_idx, val in enumerate(values, start=1):
                    cell = ws.cell(row=current_row, column=col_idx, value=val)
                    apply_fill(cell, COLOR_GREEN_CALL)
                    cell.alignment = Alignment(horizontal='right' if col_idx > 2 else 'left')
                    if col_idx in (6, 9, 10) and isinstance(val, float):
                        cell.number_format = '#,##0.00'
                    if col_idx == 11 and isinstance(val, float):
                        cell.number_format = '0.00%'
                current_row += 1

        # Trennzeile zwischen Symbolen (nicht nach dem letzten Symbol)
        if sym_idx < len(all_syms_2) - 1:
            for col in range(1, NUM_SEC2_COLS + 1):
                cell = ws.cell(row=current_row, column=col)
                apply_fill(cell, COLOR_SEPARATOR)
                cell.border = thin_border()
            current_row += 1

    # --- Spaltenbreiten automatisch anpassen ---
    max_col = max(NUM_CSP_COLS, NUM_SEC2_COLS)
    for col in range(1, max_col + 1):
        max_width = 10
        col_letter = get_column_letter(col)
        for row_cells in ws.iter_rows(min_col=col, max_col=col):
            for cell in row_cells:
                if cell.value and cell.data_type != 'n':
                    try:
                        cell_len = len(str(cell.value))
                        if cell_len > max_width:
                            max_width = cell_len
                    except Exception:
                        pass
        ws.column_dimensions[col_letter].width = min(max_width + 2, 40)

    # --- Datei speichern und Zusammenfassung ausgeben ---
    wb.save(OUTPUT_FILE)
    print(f"\nFertig! Excel-Datei gespeichert: {OUTPUT_FILE}")
    print(f"  CSPs:   {len(csp_rows)} Zeilen")
    print(f"  Aktien: {len(stock_map)} Symbole")
    print(f"  Calls:  {sum(len(v) for v in call_map.values())} Zeilen")


if __name__ == '__main__':
    main()
