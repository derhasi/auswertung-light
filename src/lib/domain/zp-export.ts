/**
 * Ausgabe im ZP-Format für den Ergebnis- und Statistikdienst von zugspitzpokal.de.
 * Spaltenaufbau wie im alten Blatt „zp_output" (21 Spalten, Dezimalpunkt, ohne Kopfzeile):
 *
 *  1 Veranstaltungs-ID (ZP)   2 Lizenz          3 Verein           4 Startnummer
 *  5 Klasse (Ziffer)          6 gewertet (1/0)  7–9 Training F1/F2/Zeit
 * 10–12 Lauf 1 F1/F2/Zeit    13–15 Lauf 2 F1/F2/Zeit              16 Gesamt
 * 17 Platz („niW" außer Wertung)                18 Punkte          19 Sportabzeichenpunkte
 *    (DNS/DSQ: gewertet = 0, Platz leer, Zeit des betroffenen Laufs 0)
 * 20 Name („Nachname, Vorname")                 21 „\N"
 */
import { stringifyCsv } from './csv';
import { anzeigeName, type Klasse, type LaufEingabe } from './typen';
import type { WertungsZeile } from './wertung';
import { punktZahl, runde } from './zahlen';

/** „Klasse 3" / „K3" → „3" (Ziffern am Ende), sonst letztes Zeichen wie im Original. */
export function klassenZiffer(klasse: Pick<Klasse, 'kuerzel' | 'name'>): string {
	const text = (klasse.kuerzel || klasse.name).trim();
	return /(\d+)$/.exec(text)?.[1] ?? text.slice(-1);
}

function laufSpalten(lauf: LaufEingabe | undefined): string[] {
	if (!lauf) return ['', '', '0'];
	return [String(lauf.fehler1 ?? ''), String(lauf.fehler2 ?? ''), punktZahl(runde(lauf.zeit ?? 0))];
}

export function zpExportZeilen(
	zpId: string,
	klassen: readonly { klasse: Klasse; zeilen: readonly WertungsZeile[] }[]
): string[][] {
	const ausgabe: string[][] = [];
	for (const { klasse, zeilen } of klassen) {
		for (const z of zeilen) {
			const s = z.starter;
			ausgabe.push([
				zpId,
				s.lizenz,
				s.verein,
				String(s.startnummer),
				klassenZiffer(klasse),
				s.ausserWertung || z.status === 'nicht-gestartet' || z.status === 'disqualifiziert' ? '0' : '1',
				...laufSpalten(s.laeufe[0]),
				...laufSpalten(s.laeufe[1]),
				...laufSpalten(s.laeufe[2]),
				punktZahl(z.gesamt ?? 0),
				s.ausserWertung ? 'niW' : z.platz === null ? '' : String(z.platz),
				punktZahl(z.punkte),
				punktZahl(z.sportabzeichen),
				anzeigeName(s),
				'\\N'
			]);
		}
	}
	return ausgabe;
}

export function zpExportCsv(zpId: string, klassen: readonly { klasse: Klasse; zeilen: readonly WertungsZeile[] }[]): string {
	return stringifyCsv(zpExportZeilen(zpId, klassen));
}
