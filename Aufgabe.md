# Schreibe ein Python Programm um meine IB Konten über die Trader Workstation auszulesen und führen Analysen aus, schreibe die Ergebnisse in eine Excel Datei

* Das Programm soll in einer grafischen Oberfläche laufen, benutze hierfür TKinker
* Alle Marktdateninformation hole dir über die API der Trader Workstation

1) Lies meine IB Konten aus und schreibe diese auf dem Hauptfenster des Programms in tabellarischer Form:
    * Folgende Spalten: Ticker Symbol, Genaue Bezeichnung, Aktueller Kurs, Strike Preis, Restlaufzeit, Verkaufspreis, EUR bzw. USD
    * Sortiere wir folgt: CSPs, Stocks+zugehörige Calls
    * Berechne bei den CSPs bzw. Calls die Restrendite falls ein Gewinn vorliegt: 
        (365/DTE)*(Erhaltene Prämie / Strike Preis)
    * Füge eine Spalte Restrendite hinzu
    * Sortiere die Aktien in zwei Gruppen: EUR und USD
    * Bestimme den noch freien Cash (also z.B. nicht für CSPs benötigt) für das EUR Guthaben als auch das USD Guthaben. Gib auch das EUR und USD Guthaben aus.

2) Erstelle eine auführliche Dokumentation im Sourcecode
    * Erkläre die Einrichtung und den Aufruf des Programms mit .venv

3) Ergänze einen Knopf für die "Auswahl CSP"

3.1) Bei Drücken des Knopfes öffne ein kleines Fenster: " Bitte Ticker eingeben"
    Ticker einlesen und prüfen ob bei IB eine Aktie zugeordnet werden kann, orientiere dich dabei an Aktien in meiner Peter Sammlung. Falls unklar, frage nach.
3.2) Lade aus IB die Werte für CSPs für einen Zeitraum von bis zu 60 Tagen, 
* Schreibe oben Links ins Fester das Ticker Sympbol, den Namen der Aktie sowie den aktuellen Kurs
* Liste die zu erwartende Rendite auf
* Lade für jede Woche die CSPs 
* Frage die Preise für Put verkaufen für jede Woche ab (heute ist der 10.3.) ... also z.Beispiel 13.3., 20.3. , 27.3, 3.4, 10,4 ... bis 8 Wochen voll sind
* Ergänze eine Spalte wo die Differenz zum Aktuellen Kurs in Prozent angegeben ist
	

4) Lade das Projekt in mein Gitbub hoch, die Github Zugangsdaten liegen in der Datei Github_Token.txt
    * Lade auch Aufgabe.md mit hoch
    * Github_Token.txt wird nicht hochgeladen, setze es auf die Liste der zu ignorierenden Dateien



    
    

    
