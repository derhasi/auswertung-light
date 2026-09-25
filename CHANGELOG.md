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

Die Änderungen der Excel-Versionen bis v0.23 stehen in [legacy/CHANGELOG.md](legacy/CHANGELOG.md).
