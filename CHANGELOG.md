# Versionsänderungen

## v2.0.0 (unveröffentlicht)

Komplette Neuentwicklung als Desktop-App (Tauri 2, Svelte 5, TypeScript, SQLite).

* Alle Funktionen der Excel-Version: Fahrerdatenbank mit ZP-Import, Nennung, Erfassung von Training und zwei
  Wertungsläufen mit zwei Fehlerarten, Ergebnisliste, Punkte, Sportabzeichenpunkte, Rookie-Markierung,
  Mannschaftswertung (Regel 2011), Zeitimport, Logos, ZP-Export
* Neu: beliebig viele Veranstaltungen und frei konfigurierbare Klassen
* Neu: Zeitmessungs-Datei (CSV/XLSX) wird beobachtet, neue Zeiten erscheinen sofort
* Neu: Druckansichten und PDF für Startliste, Ergebnisliste, Mannschaftswertung und Urkunden
* Neu: Sicherung und Übertragung von Veranstaltungen (JSON), Ergebnis-Export für Excel
* Korrektur: Klasse 6 zählt zur Mannschaftswertung
* Korrektur: Fahrer ohne vollständige Wertungsläufe werden nicht mehr vorne einsortiert

Änderungen aus dem ersten Nutzertest:

* Lizenznummern dürfen Buchstaben, Ziffern sowie `-`, `/` und `_` enthalten
* Erfasste Läufe können korrigiert werden; eine Begründung ist Pflicht und wird im Änderungsprotokoll festgehalten
* Läufe können als DNS (nicht gestartet) oder DSQ (disqualifiziert) gekennzeichnet werden, mit Pflichtkommentar
* Nach dem Speichern springt die Erfassung zum nächsten Start (Zweierpaare: Training, Wertung 1 – danach alle Wertung 2)
* Nennlisten (CSV/Excel) können je Klasse importiert werden; neue Fahrer werden automatisch angelegt, Abweichungen im Seitenvergleich zusammengeführt
* Fahrerdaten werden versioniert, Nennungen sind mit der verwendeten Version verknüpft; gleiche Lizenzen für verschiedene Fahrer sind möglich

Die Änderungen der Excel-Versionen bis v0.23 stehen in [legacy/CHANGELOG.md](legacy/CHANGELOG.md).
