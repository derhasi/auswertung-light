/**
 * Mannschaftswertung (Regel seit Saison 2011):
 * Pro Verein zählen die besten N Punktergebnisse (Standard 6) aller Klassen,
 * die für die Mannschaftswertung aktiviert sind. Gibt es auf dem letzten
 * zählenden Platz Punktgleichheit, werden die Namen mit „&" verbunden.
 *
 * Abweichend von Excel werden Vereinsnamen ohne Beachtung von Groß-/Kleinschreibung
 * und mehrfachen Leerzeichen zusammengeführt.
 */
import { runde } from './zahlen';
import { anzeigeName, type Klasse } from './typen';
import type { WertungsZeile } from './wertung';

export interface MannschaftsErgebnis {
	fahrer: string[];
	klasse: string;
	punkte: number;
}

export interface MannschaftsZeile {
	verein: string;
	platz: number;
	summe: number;
	ergebnisse: MannschaftsErgebnis[];
}

export function vereinsSchluessel(verein: string): string {
	return verein.trim().replace(/\s+/g, ' ').toLocaleLowerCase('de-DE');
}

export function mannschaftsWertung(
	klassen: readonly { klasse: Klasse; zeilen: readonly WertungsZeile[] }[],
	anzahl: number
): MannschaftsZeile[] {
	const vereine = new Map<string, { name: string; eintraege: MannschaftsErgebnis[] }>();

	for (const { klasse, zeilen } of klassen) {
		if (!klasse.inMannschaft) continue;
		for (const z of zeilen) {
			const verein = z.starter.verein.trim().replace(/\s+/g, ' ');
			if (!verein || z.punkte <= 0) continue;
			const schluessel = vereinsSchluessel(verein);
			if (!vereine.has(schluessel)) vereine.set(schluessel, { name: verein, eintraege: [] });
			vereine.get(schluessel)!.eintraege.push({
				fahrer: [anzeigeName(z.starter)],
				klasse: klasse.kuerzel || klasse.name,
				punkte: z.punkte
			});
		}
	}

	const ergebnis: Omit<MannschaftsZeile, 'platz'>[] = [];
	for (const { name, eintraege } of vereine.values()) {
		eintraege.sort((a, b) => b.punkte - a.punkte);
		const zaehlend = eintraege.slice(0, anzahl).map((e) => ({ ...e, fahrer: [...e.fahrer] }));
		const letzte = zaehlend[anzahl - 1];
		if (letzte) {
			// Punktgleiche Fahrer auf dem letzten zählenden Platz gemeinsam aufführen.
			for (const weitere of eintraege.slice(anzahl)) {
				if (weitere.punkte !== letzte.punkte) break;
				letzte.fahrer.push(...weitere.fahrer);
				if (!letzte.klasse.split(' & ').includes(weitere.klasse)) letzte.klasse += ` & ${weitere.klasse}`;
			}
		}
		const summe = runde(zaehlend.reduce((acc, e) => acc + e.punkte, 0));
		if (summe > 0) ergebnis.push({ verein: name, summe, ergebnisse: zaehlend });
	}

	ergebnis.sort((a, b) => b.summe - a.summe || a.verein.localeCompare(b.verein, 'de'));
	const zeilen: MannschaftsZeile[] = [];
	ergebnis.forEach((e, i) => {
		const vorher = zeilen[i - 1];
		zeilen.push({ ...e, platz: vorher && vorher.summe === e.summe ? vorher.platz : i + 1 });
	});
	return zeilen;
}
