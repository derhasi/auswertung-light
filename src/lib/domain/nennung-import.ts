/**
 * Nennungs-Import: ordnet die Zeilen einer Nennliste der Fahrerdatenbank zu.
 *
 * - Ein Konflikt entsteht, wenn es in der Datenbank einen Fahrer mit gleicher Lizenz
 *   oder mit gleichem Vor- und Nachnamen gibt und sich weitere Felder unterscheiden.
 * - Jedes abweichende Feld muss ausdrücklich entschieden werden (Datenbank oder Nennliste),
 *   bevor importiert werden kann. Beim Zusammenführen bleibt der alte Stand als Version erhalten.
 * - Nur bei Namensgleichheit (andere oder fehlende Lizenz) kann die Zeile stattdessen als
 *   anderer Fahrer neu angelegt werden – Lizenzen sind eindeutig.
 * - Fahrer ohne Treffer werden automatisch in die Datenbank übernommen.
 */
import { fahrerUnterschiede, type FahrerDaten, type NennungsZeile, type VergleichsFeld } from './fahrer-import';
import type { Klasse, Starter } from './typen';

export type ImportArt = 'neu' | 'bekannt' | 'konflikt' | 'bereits-gemeldet';
export type KonfliktLoesung = 'zusammenfuehren' | 'neuer-fahrer';
/** Wodurch der Datenbankeintrag gefunden wurde. */
export type Treffer = 'lizenz' | 'name';

export interface DbFahrer extends FahrerDaten {
	id: number;
}

export interface ImportPosten<F extends DbFahrer = DbFahrer> {
	zeile: NennungsZeile;
	art: ImportArt;
	treffer: Treffer | null;
	/** Mögliche Datenbankeinträge (gleiche Lizenz bzw. gleicher Name). */
	kandidaten: F[];
	/** Datensatz, mit dem verglichen wird. */
	bestand: F | null;
	unterschiede: VergleichsFeld[];
	loesung: KonfliktLoesung;
	/** Entscheidung je abweichendem Feld – ohne Vorauswahl. */
	auswahl: Partial<Record<VergleichsFeld, 'bestand' | 'import'>>;
	klasseId: number | null;
	startnummer: number | null;
	uebernehmen: boolean;
	hinweis: string;
	/** Hinweis zur Startnummer (z. B. gewünschte Nummer vergeben). */
	nummerHinweis: string;
}

const nameSchluessel = (f: Pick<FahrerDaten, 'nachname' | 'vorname'>) =>
	`${f.nachname}|${f.vorname}`.trim().replace(/\s+/g, ' ').toLocaleLowerCase('de-DE');

export function klasseZuordnen(klassen: readonly Klasse[], text: string): number | null {
	const t = text.trim().toLowerCase();
	if (!t) return null;
	return klassen.find((k) => k.kuerzel.toLowerCase() === t || k.name.toLowerCase() === t)?.id ?? null;
}

/** Vergleicht mit einem Kandidaten und setzt Art und Unterschiede (ohne Vorauswahl). */
export function mitBestandVergleichen<F extends DbFahrer>(posten: ImportPosten<F>, bestand: F | null): ImportPosten<F> {
	const unterschiede = bestand ? fahrerUnterschiede(bestand, posten.zeile) : [];
	return {
		...posten,
		bestand,
		unterschiede,
		art: posten.art === 'bereits-gemeldet' ? posten.art : !bestand ? 'neu' : unterschiede.length ? 'konflikt' : 'bekannt',
		auswahl: {},
		loesung: 'zusammenfuehren'
	};
}

/** Ist der Posten importierbar (Konflikt vollständig entschieden)? */
export function konfliktGeloest(p: ImportPosten): boolean {
	if (p.art !== 'konflikt') return true;
	if (p.loesung === 'neuer-fahrer') return p.treffer === 'name';
	return p.unterschiede.every((feld) => p.auswahl[feld] !== undefined);
}

export function offeneFelder(p: ImportPosten): VergleichsFeld[] {
	if (p.art !== 'konflikt' || p.loesung === 'neuer-fahrer') return [];
	return p.unterschiede.filter((feld) => p.auswahl[feld] === undefined);
}

export function importPlanen<F extends DbFahrer>(
	zeilen: readonly NennungsZeile[],
	datenbank: readonly F[],
	vorhandeneStarter: readonly Starter[],
	klassen: readonly Klasse[],
	zielKlasseId: number | null
): ImportPosten<F>[] {
	const nachLizenz = new Map<string, F>();
	const nachName = new Map<string, F[]>();
	for (const f of datenbank) {
		if (f.lizenz) nachLizenz.set(f.lizenz, f);
		nachName.set(nameSchluessel(f), [...(nachName.get(nameSchluessel(f)) ?? []), f]);
	}
	const gemeldeteLizenzen = new Set(vorhandeneStarter.map((s) => s.lizenz).filter(Boolean));
	const gemeldeteIds = new Set(vorhandeneStarter.map((s) => s.fahrerId).filter((id): id is number => typeof id === 'number'));

	const posten = zeilen.map((zeile) => {
		const perLizenz = zeile.lizenz ? nachLizenz.get(zeile.lizenz) : undefined;
		const perName = nachName.get(nameSchluessel(zeile)) ?? [];
		const treffer: Treffer | null = perLizenz ? 'lizenz' : perName.length ? 'name' : null;
		const kandidaten = perLizenz ? [perLizenz] : perName;
		// Bei mehreren Namensgleichen: den mit den wenigsten Abweichungen vergleichen
		const bestand = [...kandidaten].sort((a, b) => fahrerUnterschiede(a, zeile).length - fahrerUnterschiede(b, zeile).length)[0] ?? null;
		const klasseId = zielKlasseId ?? klasseZuordnen(klassen, zeile.klasse);
		const gemeldet = (zeile.lizenz && gemeldeteLizenzen.has(zeile.lizenz)) || (bestand && gemeldeteIds.has(bestand.id));
		const basis: ImportPosten<F> = {
			zeile,
			art: gemeldet ? 'bereits-gemeldet' : 'neu',
			treffer,
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

/** Datensatz, der beim Zusammenführen entsteht (nicht entschiedene Felder behalten den Datenbankwert). */
export function zusammengefuehrt<F extends DbFahrer>(p: ImportPosten<F>): FahrerDaten {
	if (!p.bestand) return p.zeile;
	const { id: _id, ...bestand } = p.bestand;
	const ergebnis: FahrerDaten = { ...(bestand as FahrerDaten) };
	for (const feld of p.unterschiede) {
		if (p.auswahl[feld] === 'import') ergebnis[feld] = p.zeile[feld];
	}
	// Leere Felder im Bestand aus dem Import ergänzen
	for (const feld of ['plz', 'ort', 'verein', 'geburtsdatum', 'vorname'] as const) {
		if (!ergebnis[feld] && p.zeile[feld]) ergebnis[feld] = p.zeile[feld];
	}
	if (ergebnis.rookieJahr === null && p.zeile.rookieJahr !== null) ergebnis.rookieJahr = p.zeile.rookieJahr;
	return ergebnis;
}
