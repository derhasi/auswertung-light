# Auswertung Light

**English**: Documentation of all other parts of Code will be held in German, as this tool is intended to serve German Kart-Slalom Clubs. Further information can be fetched at zugspitzpokal.de

Auswertung-Light ist eine Excel-Arbeitsmappe die im Laufe der Jahre durch Implementierung von VBA-Scripten zu einem umfangreicheren Auswertprogramm für den Kart-Slalom geworden ist.

## Features

Das Programm enthält folgende Features:
* automatische Eintragung der Fahrerdaten über die Eingabe der Fahrerlizenz aus dazugehöriger Fahrerdatentabelle
* einfache Erstellung der Ergebnisliste durch Knopfdruck
* einfache Sortierung nach Startnummern durch Knopfdruck
* einfach Erstellung der Mannschaftswertung durch Knopfdruck (NEU ab v021 mit Mannschaftswertung für Saison 2011)
* Anzeige von bis zu 2 Fehlerspalten plus Zeit, um z.B. zwischen Pylonen-(2sec) und Torfehlern(10sec) unterscheiden zu können
* Anzeige der Adressdaten (wie es der ADAC verlangt)
* Anzeige der ADAC-Sportabzeichenpunkte
* Ausgabe im ZP-Format (für den zugspitzpokal.de Ergebnis- und Statistikdienst)
* Import der Zugspitzpokal-Fahrerdatenliste
* Angepasste Menüleiste
* Hilfe zum anpassen der Zahlenformatierung (von Michael Steinhoff)
* Markierung der Rookies
* Einfügen von Header-Logos automatisiert
* Import einer Zeit aus externer Excelmapper "Zeitenimport"
* Sortierung der Datentabelle nach Lizenz, Name oder Klasse
* --GEPLANT Datenexport mit Bildimplementierungsmöglickeit (initiiert vom MC Dießen)--

![Screenshot menu](http://www.zugspitzpokal.de/sites/default/files/auswertung_light_v017_menubar.png)

### Integrierte Fahrerdatenbank

Ab der Version 0.17 wurde eine Fahrerdaten-Importfunktion eingebaut. Diese liest die Zugspitzpokal-Fahrerliste ein, welche automatisch auf zugspitzpokal.de generiert wird. 
Diese Fahrerliste ist als CSV-Download für angemeldete Nutzer (Datenpfleger) stets aktuell* verfügbar.

_*Sie wird von den Verantwortlichen der Vereine gepflegt_

## Team:

Hauptprogramm: [Johannes Haseitl](http://derhasi.de)
Weitere Hilfe: Michael Steinhoff(Zahlenformat: deutsch/englisch), Dieter Schweingruber (Testing, MC Dießen)

## Fragen und Antworten

Bei Fragen zum Programm wendet euch bitte an das Team oder hinterlasst einen Kommentar im Issue-Bereich auf https://github.com/derhasi/auswertung-light/issues.

## Versionsänderungen

Unter https://github.com/derhasi/auswertung-light/blob/master/CHANGELOG.md ist die Liste der Ängerungen ein zu sehen.
