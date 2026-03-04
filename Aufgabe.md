# Schreibe ein Python Programm um meine IB Konten über die Trader Workstation auszulesen und führen Analysen aus, schreibe die Ergebnisse in eine Excel Datei

1) Lies meine IB Konten aus und schreibe diese in eine Excel Tabelle
    * Folgende Spalten in der Excel Tabelle: Ticker Symbol, Genaue Bezeichnung, Aktueller Kurs, Strike Preis, Restlaufzeit, Verkaufspreis, EUR bzw. USD
    * Sortiere wir folgt: CSPs, Stocks+zugehörige Calls
    * Berechne bei den CSPs bzw. Calls die Restrendite falls ein Gewinn vorliegt: 
        (365/DTE)*(Erhaltene Prämie / Strike Preis)
    * Füge eine Spalte Restrendite hinzu
    * Sortiere die Aktien in zwei Gruppen: EUR und USD
    * Bestimme den noch freien Cash (also z.B. nicht für CSPs benötigt) für das EUR Guthaben als auch das USD Guthaben. Gib auch das EUR und USD Guthaben aus.

2) Erstelle eine auführliche Dokumentation im Sourcecode
    * Erkläre die Einrichtung und den Aufruf des Programms mit .venv

3) Lade das Projekt in mein Gitbub hoch, die Github Zugangsdaten liegen in der Datei Github_Token.txt
    * Lade auch Aufgabe.md mit hoch
    * Github_Token.txt wird nicht hochgeladen, setze es auf die Liste der zu ignorierenden Dateien

    
    

    
