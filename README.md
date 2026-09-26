# Auswertung Light 2

**English**: Auswertung Light is a desktop app for scoring German kart slalom events (Zugspitzpokal). The documentation is in German, as the tool is intended for German kart slalom clubs.

Auswertung Light ist ein Auswertprogramm für den **Kart-Slalom**. Version 2 ist eine Neuentwicklung der bisherigen
Excel-Arbeitsmappe als **Desktop-App für Windows, macOS und Linux**. Sie arbeitet vollständig offline, sodass
am Veranstaltungsort kein Internet nötig ist.

![Erfassung](docs/screenshots/erfassung.png)

## Funktionen

| Bereich | Was die App kann |
|---|---|
| **Fahrerdatenbank** | Import der Zugspitzpokal-Fahrerliste (CSV, UTF-8 oder Windows-1252). Vorhandene Fahrer werden abgeglichen statt überschrieben. Jede Änderung wird als Version gespeichert; jede Nennung ist mit der Version verknüpft, mit der gemeldet wurde. Lizenzen sind eindeutig und dürfen Buchstaben, Ziffern und `-` `/` `_` enthalten. Suche, Sortierung, Bearbeiten, CSV-Export, Übersicht der Starts je Fahrer. |
| **Veranstaltungen** | Beliebig viele Veranstaltungen. Klassen, Strafsekunden, Logos, Ausrichter und Zeitmessung werden von der letzten Veranstaltung übernommen. |
| **Klassen** | Frei konfigurierbar (Standard: Klasse 1–6), Reihenfolge änderbar, je Klasse einstellbar, ob sie zur Mannschaftswertung zählt. |
| **Nennung** | Fahrer per Suche (Name, Lizenz, Verein) mit der Tastatur nennen oder eine Nennliste (CSV/Excel, optional mit Startnummern) je Klasse importieren. Unbekannte Fahrer werden automatisch in die Datenbank übernommen. Gibt es bereits einen Fahrer mit gleicher Lizenz oder gleichem Vor- und Nachnamen, werden die Datensätze nebeneinander angezeigt; jedes abweichende Feld muss auf den Datenbank- oder Nennlisten-Wert festgelegt werden. Bei Namensgleichheit kann die Zeile stattdessen als anderer Fahrer angelegt werden. Kennzeichnung „außer Wertung" (niW). |
| **Erfassung** | Eingabemaske ganz für die Tastatur: Fehler 1 ↵, Fehler 2 ↵, Zeit ↵. Nach dem Speichern springt die Maske zum nächsten Start der Reihenfolge – Klasse für Klasse: je zwei Fahrer Training und Wertung 1, danach alle Fahrer der Klasse Wertung 2. Nach dem letzten Lauf einer Klasse folgt ein Zwischenschritt mit „Ergebnis anzeigen“ und „Zur nächsten Klasse“. DNS/DSQ mit Pflichtkommentar. Bereits erfasste Läufe können korrigiert werden – mit Pflicht-Begründung und Änderungsprotokoll, auch nachträglich aus der Ergebnisliste. Zeiten als `32,45` oder `1:02,34`. |
| **Zeitmessung** | Übernahme von Zeiten aus der CSV- oder Excel-Datei der Zeitmessanlage. Die Datei wird beobachtet, neue Zeiten erscheinen sofort. Übernahme per Klick oder Strg + T, bereits zugeordnete Zeiten werden markiert. |
| **Ergebnisse** | Live berechnete Ergebnislisten je Klasse mit Laufzeiten, Strafen, Gesamtzeit, Punkten und ADAC-Sportabzeichenpunkten. Rookies werden markiert, Adressen lassen sich einblenden. |
| **Mannschaftswertung** | Die besten 6 Ergebnisse (einstellbar) je Verein über alle Klassen. Bei Punktgleichstand auf dem letzten zählenden Platz werden die Namen mit „&" verbunden. |
| **Drucken / PDF** | Startliste mit Notizfeldern, Ergebnisliste, Mannschaftswertung und Urkunden, jeweils mit Vereinslogos im Kopf. |
| **Export** | ZP-Format (`zp_output.csv`) für den Ergebnisdienst von zugspitzpokal.de, Ergebnisse als CSV für Excel, Sicherung und Übertragung einer Veranstaltung als JSON. |

<details>
<summary>Weitere Screenshots</summary>

![Nennung](docs/screenshots/nennung.png)
![Ergebnisse](docs/screenshots/ergebnisse.png)
![Mannschaftswertung](docs/screenshots/mannschaft.png)

</details>

## Wertungsregeln

Die Rechenregeln stammen aus den Formeln der Excel-Version (`src/lib/domain/wertung.ts`, `mannschaft.ts`):

- **Laufergebnis** = Zeit + Fehler 1 × Strafe 1 + Fehler 2 × Strafe 2 (Standard: Pylone 2 s, Tor 10 s), auf 1/100 gerundet
- **Gesamt** = Wertungslauf 1 + Wertungslauf 2. Das Training zählt nicht.
- **Reihenfolge:** Gesamtzeit, bei Gleichstand der bessere Einzellauf. Sind beide gleich, gibt es den gleichen Platz (1, 1, 3 …).
- **Punkte** = (Teilnehmer − Platz) × 10 / Teilnehmer + 1. Alle gemeldeten Fahrer der Klasse zählen als Teilnehmer.
- **Sportabzeichenpunkte:** Platz 1 = 6, Platz 2–10 = (12 − Platz) / 2, ab Platz 11 = 0,5
- **Mannschaft:** Summe der besten N Punktergebnisse je Verein
- **DNS / DSQ:** Ist ein Wertungslauf als DNS (nicht gestartet) oder DSQ (disqualifiziert) gekennzeichnet, erhält der Fahrer keinen Platz und keine Punkte. Im ZP-Export steht er als „nicht gewertet“ (0) ohne Platz.

### Unterschiede zur Excel-Version

- Fahrer ohne beide Wertungsläufe werden als „unvollständig" ohne Platz geführt. Bisher rutschten sie mit 0 s nach vorne.
- Klasse 6 zählt jetzt zur Mannschaftswertung. In v0.23 fehlte sie dort.
- Vereinsnamen werden ohne Beachtung von Groß-/Kleinschreibung und doppelten Leerzeichen zusammengeführt.
- Startnummern sind je Veranstaltung eindeutig.
- Lizenznummern bleiben Text, führende Nullen gehen also nicht verloren.
- Nennungen speichern eine Kopie der Fahrerdaten. Spätere Änderungen an der Fahrerdatenbank verändern alte Ergebnisse nicht.

## ZP-Format

`zp_output.csv` hat denselben Aufbau wie das frühere Blatt „zp_output": 21 Spalten, kommagetrennt,
Dezimalpunkt, Windows-1252, ohne Kopfzeile.

`Veranstaltungs-ID, Lizenz, Verein, Startnummer, Klasse, gewertet (1/0), Training F1/F2/Zeit, Lauf 1 F1/F2/Zeit, Lauf 2 F1/F2/Zeit, Gesamt, Platz (bzw. niW), Punkte, Sportabzeichenpunkte, Name, \N`

Die Veranstaltungs-ID wird unter *Drucken & Export* oder in den Einstellungen der Veranstaltung eingetragen.

## Installation

Installationsdateien für Windows (`.msi`/`.exe`), macOS (`.dmg`) und Linux (`.deb`/`.rpm`/`.AppImage`) baut
der Workflow *Build* in GitHub Actions:

- bei jedem Push auf `main` als Artefakte am jeweiligen Workflow-Lauf (Actions → Build → Lauf → *Artifacts*)
- bei einem Versions-Tag (`v2.0.0`) zusätzlich als Entwurf unter *Releases*
- oder per Hand: *Actions → Build → Run workflow*, Branch wählen und unter „Release-Tag" z. B. `v2.0.0`
  eintragen. Es entsteht ein Release-Entwurf für den aktuellen Stand des Branches; den Tag legt GitHub an,
  sobald der Entwurf veröffentlicht wird.

### Unsignierte Pakete trotzdem starten

Die Installationspakete sind nicht mit einem kostenpflichtigen Entwickler-Zertifikat signiert. Windows und macOS
warnen deshalb beim ersten Start. Die App funktioniert trotzdem ganz normal, die Freigabe ist nur einmal nötig.
Pakete bitte nur von der [Releases-Seite dieses Repositorys](https://github.com/derhasi/auswertung-light/releases)
herunterladen.

#### Windows

1. **Beim Herunterladen** (Edge/Chrome): Meldet der Browser „wird nicht häufig heruntergeladen“ oder
   „könnte schädlich sein“, im Download-Bereich über **…** → **Behalten** (ggf. **Trotzdem behalten**) bestätigen.
2. **Beim Start des Installers** (`…_x64-setup.exe` oder `…_x64_de-DE.msi`) erscheint
   „Der Computer wurde durch Windows geschützt“ (SmartScreen):
   auf **Weitere Informationen** klicken, dann **Trotzdem ausführen**.
3. Falls die Schaltfläche fehlt: Rechtsklick auf die heruntergeladene Datei → **Eigenschaften** → unten bei
   „Sicherheit“ **Zulassen** anhaken → **OK**, danach erneut starten.

Nach der Installation startet die App ohne weitere Warnung.

#### macOS

1. Die `.dmg`-Datei öffnen und **Auswertung Light** in den Ordner **Programme** ziehen.
2. Die App im Ordner **Programme** öffnen. macOS meldet, dass die App nicht geöffnet werden kann, weil der
   Entwickler nicht verifiziert werden kann. Mit **Fertig** bzw. **OK** schließen (nicht „In den Papierkorb legen“).
3. **Systemeinstellungen** → **Datenschutz & Sicherheit** öffnen, nach unten zum Abschnitt „Sicherheit“ scrollen
   und bei „Auswertung Light wurde blockiert …“ auf **Dennoch öffnen** klicken. Mit dem Passwort bzw. Touch ID
   bestätigen und im folgenden Dialog nochmals **Dennoch öffnen** wählen.
   (Bis macOS 14 geht es auch kürzer: Rechtsklick auf die App → **Öffnen** → **Öffnen**.)

Meldet macOS stattdessen **„Auswertung Light“ ist beschädigt und kann nicht geöffnet werden**, ist die App nicht
defekt – macOS blockiert so unsignierte Apps aus dem Internet. Dann im **Terminal** einmalig die
Download-Markierung entfernen und die App erneut öffnen:

```bash
xattr -dr com.apple.quarantine "/Applications/Auswertung Light.app"
```

#### Linux

Das `.deb`- bzw. `.rpm`-Paket wie gewohnt installieren. Ein `.AppImage` vorher ausführbar machen
(`chmod +x "Auswertung Light_2.0.0_amd64.AppImage"`) und dann starten.

### Daten

Die Daten liegen in der SQLite-Datei `auswertung-light.db` im App-Datenverzeichnis, unter Windows
`%APPDATA%\de.zugspitzpokal.auswertung-light`. Für einen Rechnerwechsel am besten die Veranstaltung unter
*Drucken & Export → Veranstaltung sichern* exportieren und auf der Startseite wieder importieren.

## Entwicklung

Voraussetzungen: [Node.js](https://nodejs.org) ≥ 20, [pnpm](https://pnpm.io), [Rust](https://rustup.rs) und die
[Tauri-Systemvoraussetzungen](https://v2.tauri.app/start/prerequisites/).

```bash
pnpm install
pnpm tauri dev     # Desktop-App mit Hot Reload
pnpm dev           # nur die Oberfläche im Browser (Demo-Modus, Daten im localStorage)
pnpm test          # Unit-Tests (Wertung, Import/Export, Datenbank)
pnpm check         # Typprüfung
pnpm tauri build   # Installationspaket für das aktuelle Betriebssystem
```

### Aufbau

```
src/
├─ lib/domain/     Fachlogik als reine TypeScript-Funktionen (mit Tests)
│  ├─ wertung.ts         Laufergebnis, Rangliste, Punkte
│  ├─ mannschaft.ts      Mannschaftswertung
│  ├─ fahrer-import.ts   ZP-Fahrerliste (CSV)
│  ├─ zp-export.ts       ZP-Output
│  └─ zeitquelle.ts      Zeitmessung (CSV/XLSX)
├─ lib/db/         SQLite-Schema (Drizzle ORM), Migrationen, Repository
│                  Desktop: tauri-plugin-sql · Browser/Tests: sql.js
├─ lib/stores/     Zustand einer geöffneten Veranstaltung (Svelte 5 Runes)
├─ lib/components/ Ergebnisliste, Mannschaftsliste, Fahrersuche, Druckkopf
└─ routes/         Seiten (SvelteKit, statisch als SPA ausgeliefert)
src-tauri/         Desktop-Hülle (Rust): Plugins für SQL, Dialoge, Dateien, Drucken
legacy/            Die bisherige Excel-Version (v0.23) samt exportiertem VBA-Code
```

Schemaänderungen: `src/lib/db/schema.ts` anpassen und `pnpm db:generate` ausführen. Die neue Migration wird beim
nächsten Start automatisch angewendet.

## Team

Hauptprogramm: [Johannes Haseitl](http://derhasi.de)
Weitere Hilfe (Excel-Version): Michael Steinhoff (Zahlenformat deutsch/englisch), Dieter Schweingruber (Testing, MC Dießen)

Fragen und Fehler bitte unter https://github.com/derhasi/auswertung-light/issues melden.
Die Versionsänderungen stehen im [CHANGELOG](CHANGELOG.md).
