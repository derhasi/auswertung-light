/**
 * Startreihenfolge am Veranstaltungstag – Klasse für Klasse:
 * Innerhalb einer Klasse starten die Fahrer in Zweierpaaren, erst beide Training,
 * dann beide Wertungslauf 1. Sind alle Paare der Klasse durch, fährt jeder Fahrer
 * der Klasse (in derselben Reihenfolge) Wertungslauf 2. Danach folgt die nächste Klasse.
 *
 *   Klasse 1: 1 T, 2 T, 1 W1, 2 W1, 3 T, 3 W1, 1 W2, 2 W2, 3 W2 · Klasse 2: …
 *
 * Die Klassen folgen ihrer eingestellten Reihenfolge, die Fahrer der Startnummer.
 */
import { laufErfasst, type LaufNr, type Starter } from './typen';

export interface StartPlatz {
	starter: Starter;
	lauf: LaufNr;
}

export function sortierteStarter(
	starter: readonly Starter[],
	klassenPosition: ReadonlyMap<number, number> = new Map()
): Starter[] {
	return [...starter].sort(
		(a, b) =>
			(klassenPosition.get(a.klasseId) ?? 0) - (klassenPosition.get(b.klasseId) ?? 0) ||
			a.startnummer - b.startnummer
	);
}

export function startReihenfolge(
	starter: readonly Starter[],
	klassenPosition: ReadonlyMap<number, number> = new Map()
): StartPlatz[] {
	const liste = sortierteStarter(starter, klassenPosition);
	const reihenfolge: StartPlatz[] = [];
	let i = 0;
	while (i < liste.length) {
		// Alle Fahrer derselben Klasse
		const klasseId = liste[i].klasseId;
		let ende = i;
		while (ende < liste.length && liste[ende].klasseId === klasseId) ende++;
		const klasse = liste.slice(i, ende);
		for (let j = 0; j < klasse.length; j += 2) {
			const paar = klasse.slice(j, j + 2);
			for (const lauf of [0, 1] as LaufNr[]) {
				for (const s of paar) reihenfolge.push({ starter: s, lauf });
			}
		}
		for (const s of klasse) reihenfolge.push({ starter: s, lauf: 2 });
		i = ende;
	}
	return reihenfolge;
}

/** Sind alle Läufe (Training, W1, W2) aller Fahrer der Klasse erfasst? */
export function klasseKomplett(starter: readonly Starter[], klasseId: number): boolean {
	const inKlasse = starter.filter((s) => s.klasseId === klasseId);
	return inKlasse.length > 0 && inKlasse.every((s) => ([0, 1, 2] as LaufNr[]).every((nr) => laufErfasst(s.laeufe[nr])));
}

/**
 * Nächster noch nicht erfasster Start nach dem angegebenen (oder vom Anfang).
 * Ist hinter dem aktuellen Start alles erfasst, wird vorne weitergesucht.
 */
export function naechsterStart(
	reihenfolge: readonly StartPlatz[],
	nach?: { starterId: number; lauf: LaufNr } | null
): StartPlatz | null {
	const index = nach ? reihenfolge.findIndex((p) => p.starter.id === nach.starterId && p.lauf === nach.lauf) : -1;
	for (let i = 1; i <= reihenfolge.length; i++) {
		const platz = reihenfolge[(index + i + reihenfolge.length) % reihenfolge.length];
		if (!laufErfasst(platz.starter.laeufe[platz.lauf])) return platz;
	}
	return null;
}

/** Die nächsten offenen Starts ab einer Position (für die Vorschau). */
export function offeneStarts(
	reihenfolge: readonly StartPlatz[],
	ab?: { starterId: number; lauf: LaufNr } | null,
	anzahl = 8
): StartPlatz[] {
	const index = ab ? reihenfolge.findIndex((p) => p.starter.id === ab.starterId && p.lauf === ab.lauf) : -1;
	const ergebnis: StartPlatz[] = [];
	for (let i = 1; i <= reihenfolge.length && ergebnis.length < anzahl; i++) {
		const platz = reihenfolge[(index + i + reihenfolge.length) % reihenfolge.length];
		if (!laufErfasst(platz.starter.laeufe[platz.lauf])) ergebnis.push(platz);
	}
	return ergebnis;
}
