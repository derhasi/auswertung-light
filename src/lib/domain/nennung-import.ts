/**
 * Nennungs-Import: ordnet die Zeilen einer Nennliste der Fahrerdatenbank zu.
 *
 * - Unbekannte Fahrer werden automatisch in die Datenbank übernommen.
 * - Bekannte Fahrer mit identischen Daten werden direkt genannt.
 * - Weichen Name, Adresse oder Verein ab, entsteht ein Konflikt. Er wird im Dialog
 *   gelöst: zusammenführen (feldweise), bestehenden Datensatz verwenden oder als
 *   anderen Fahrer neu anlegen. Beim Zusammenführen bleibt der alte Stand als
 *   Version erhalten.
 */
import { fahrerUnterschiede, type FahrerDaten, type NennungsZeile, type VergleichsFeld } from './fahrer-import';
import type { Klasse, Starter } from './typen';

export type ImportArt = 'neu' | 'bekannt' | 'konflikt' | 'bereits-gemeldet';
export type KonfliktLoesung = 'zusammenfuehren' | 'bestehend' | 'neuer-fahrer';

export interface DbFahrer extends FahrerDaten {
	id: number;
}

export interface ImportPosten<F extends DbFahrer = DbFahrer> {
	zeile: NennungsZeile;
	art: ImportArt;
	/** Fahrer der Datenbank mit gleicher Lizenz (bzw. gleichem Namen, wenn die Lizenz fehlt). */
	kandidaten: F[];
	/** Datensatz, mit dem verglichen wird. */
	bestand: F | null;
	unterschiede: VergleichsFeld[];
	loesung: KonfliktLoesung;
	/** Beim Zusammenführen: welcher Wert je abweichendem Feld übernommen wird. */
	auswahl: Partial<Record<VergleichsFeld, 'bestand' | 'import'>>;
	klasseId: number | null;
	startnummer: number | null;
	uebernehmen: boolean;
	hinweis: string;
	/** Hinweis zur Startnummer (z. B. gewünschte Nummer vergeben). */
	nummerHinweis: string;
}

const nameSchluessel = (f: Pick<FahrerDaten, 'nachname' | 'vorname'>) =>
	`${f.nachname} ${f.vorname}`.trim().replace(/\s+/g, ' ').toLocaleLowerCase('de-DE');

export function klasseZuordnen(klassen: readonly Klasse[], text: string): number | null {
	const t = text.trim().toLowerCase();
	if (!t) return null;
	return klassen.find((k) => k.kuerzel.toLowerCase() === t || k.name.toLowerCase() === t)?.id ?? null;
}

/** Vergleicht mit einem Kandidaten und setzt Art und Standardauswahl. */
export function mitBestandVergleichen<F extends DbFahrer>(posten: ImportPosten<F>, bestand: F | null): ImportPosten<F> {
	const unterschiede = bestand ? fahrerUnterschiede(bestand, posten.zeile) : [];
	return {
		...posten,
		bestand,
		unterschiede,
		art: posten.art === 'bereits-gemeldet' ? posten.art : !bestand ? 'neu' : unterschiede.length ? 'konflikt' : 'bekannt',
		auswahl: Object.fromEntries(unterschiede.map((f) => [f, 'import'])),
		loesung: 'zusammenfuehren'
	};
}

export function importPlanen<F extends DbFahrer>(
	zeilen: readonly NennungsZeile[],
	datenbank: readonly F[],
	vorhandeneStarter: readonly Starter[],
	klassen: readonly Klasse[],
	zielKlasseId: number | null
): ImportPosten<F>[] {
	const nachLizenz = new Map<string, F[]>();
	const nachName = new Map<string, F[]>();
	for (const f of datenbank) {
		if (f.lizenz) nachLizenz.set(f.lizenz, [...(nachLizenz.get(f.lizenz) ?? []), f]);
		nachName.set(nameSchluessel(f), [...(nachName.get(nameSchluessel(f)) ?? []), f]);
	}
	const gemeldeteLizenzen = new Set(vorhandeneStarter.map((s) => s.lizenz).filter(Boolean));
	const gemeldeteIds = new Set(vorhandeneStarter.map((s) => s.fahrerId).filter((id): id is number => typeof id === 'number'));

	const posten = zeilen.map((zeile) => {
		const kandidaten = zeile.lizenz
			? (nachLizenz.get(zeile.lizenz) ?? [])
			: (nachName.get(nameSchluessel(zeile)) ?? []).filter((f) => !f.lizenz || !zeile.lizenz);
		// Bevorzugt: identischer Datensatz, sonst gleicher Name, sonst der erste
		const bestand =
			kandidaten.find((k) => fahrerUnterschiede(k, zeile).length === 0) ??
			kandidaten.find((k) => nameSchluessel(k) === nameSchluessel(zeile)) ??
			kandidaten[0] ??
			null;
		const klasseId = zielKlasseId ?? klasseZuordnen(klassen, zeile.klasse);
		const gemeldet = (zeile.lizenz && gemeldeteLizenzen.has(zeile.lizenz)) || (bestand && gemeldeteIds.has(bestand.id));
		const basis: ImportPosten<F> = {
			zeile,
			art: gemeldet ? 'bereits-gemeldet' : 'neu',
			kandidaten,
			bestand: null,
			unterschiede: [],
			loesung: 'zusammenfuehren',
			auswahl: {},
			klasseId,
			startnummer: null,
			uebernehmen: !gemeldet && klasseId !== null,
			hinweis: gemeldet ? 'Bereits gemeldet' : klasseId === null ? `Klasse „${zeile.klasse || '–'}“ unbekannt` : '',
			nummerHinweis: ''
		};
		return mitBestandVergleichen(basis, bestand);
	});
	return startnummernVergeben(posten, vorhandeneStarter);
}

/**
 * Vergibt Startnummern: Nummern aus der Datei werden übernommen, wenn sie frei sind,
 * sonst die nächste freie Nummer nach der höchsten der Klasse.
 */
export function startnummernVergeben<P extends ImportPosten>(posten: P[], vorhandeneStarter: readonly Pick<Starter, 'klasseId' | 'startnummer'>[]): P[] {
	const belegt = new Set(vorhandeneStarter.map((s) => s.startnummer));
	const hoechste = new Map<number, number>();
	for (const s of vorhandeneStarter) hoechste.set(s.klasseId, Math.max(hoechste.get(s.klasseId) ?? 0, s.startnummer));
	const aktive = posten.filter((p) => p.uebernehmen && p.klasseId !== null);
	// Zuerst die gewünschten Nummern reservieren
	const gewuenscht = new Map<P, number>();
	for (const p of aktive) {
		const nr = p.zeile.startnummer;
		if (nr && !belegt.has(nr)) {
			belegt.add(nr);
			gewuenscht.set(p, nr);
		}
	}
	return posten.map((p) => {
		if (!p.uebernehmen || p.klasseId === null) return { ...p, startnummer: null, nummerHinweis: '' };
		let nr = gewuenscht.get(p);
		let nummerHinweis = '';
		if (nr === undefined) {
			const basis = hoechste.get(p.klasseId) ?? Math.max(0, ...belegt);
			nr = basis + 1;
			while (belegt.has(nr)) nr++;
			belegt.add(nr);
			if (p.zeile.startnummer) nummerHinweis = `Nr. ${p.zeile.startnummer} ist vergeben – ${nr} zugeteilt`;
		}
		hoechste.set(p.klasseId, Math.max(hoechste.get(p.klasseId) ?? 0, nr));
		return { ...p, startnummer: nr, nummerHinweis };
	});
}

/** Datensatz, der beim Zusammenführen entsteht. */
export function zusammengefuehrt<F extends DbFahrer>(p: ImportPosten<F>): FahrerDaten {
	if (!p.bestand) return p.zeile;
	const ergebnis: FahrerDaten = { ...p.bestand };
	for (const feld of p.unterschiede) {
		if ((p.auswahl[feld] ?? 'import') === 'import') ergebnis[feld] = p.zeile[feld];
	}
	// Leere Felder im Bestand aus dem Import ergänzen
	for (const feld of ['plz', 'ort', 'verein', 'geburtsdatum', 'vorname'] as const) {
		if (!ergebnis[feld] && p.zeile[feld]) ergebnis[feld] = p.zeile[feld];
	}
	if (ergebnis.rookieJahr === null && p.zeile.rookieJahr !== null) ergebnis.rookieJahr = p.zeile.rookieJahr;
	return ergebnis;
}
