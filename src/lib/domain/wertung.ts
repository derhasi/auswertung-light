/**
 * Klassenwertung – übernimmt die Rechenregeln der alten Excel-Auswertung:
 *
 * - Laufergebnis = Fehler1 × Strafe1 + Fehler2 × Strafe2 + Zeit (auf 1/100 gerundet)
 * - Gesamt = Wertungslauf 1 + Wertungslauf 2 (Training zählt nicht)
 * - Reihenfolge: Gesamt aufsteigend, bei Gleichstand entscheidet der bessere Einzellauf
 * - Gleiche Gesamtzeit und gleicher bester Lauf ⇒ gleicher Platz (1, 1, 3 …)
 * - Punkte = (Teilnehmer − Platz) × 10 / Teilnehmer + 1 (auf 1/100 gerundet)
 * - Sportabzeichenpunkte: Platz 1 = 6, Platz 2–10 = (12 − Platz) / 2, ab Platz 11 = 0,5
 *
 * Abweichend von Excel werden Fahrer ohne vollständige Wertungsläufe nicht
 * mehr (mit 0 s) vorne einsortiert, sondern als „unvollständig" ohne Platz geführt.
 */
import { hundertstel, runde } from './zahlen';
import type { LaufEingabe, LaufNr, Regeln, Starter } from './typen';

export type WertungsStatus = 'gewertet' | 'unvollstaendig' | 'ausser-wertung';

export interface WertungsZeile {
	starter: Starter;
	/** Laufergebnisse inkl. Strafsekunden; `null` = keine Zeit erfasst. */
	ergebnisse: Record<LaufNr, number | null>;
	gesamt: number | null;
	bester: number | null;
	platz: number | null;
	status: WertungsStatus;
	punkte: number;
	sportabzeichen: number;
	rookie: boolean;
}

export function laufErgebnis(lauf: LaufEingabe | undefined, regeln: Pick<Regeln, 'strafe1' | 'strafe2'>): number | null {
	if (!lauf || lauf.zeit === null || lauf.zeit === undefined) return null;
	return runde((lauf.fehler1 || 0) * regeln.strafe1 + (lauf.fehler2 || 0) * regeln.strafe2 + lauf.zeit);
}

export function istRookie(starter: Pick<Starter, 'rookieJahr'>, veranstaltungsJahr: number | null): boolean {
	return starter.rookieJahr !== null && veranstaltungsJahr !== null && starter.rookieJahr === veranstaltungsJahr;
}

export function wertungsPunkte(platz: number, teilnehmer: number): number {
	if (teilnehmer <= 0) return 0;
	return runde(((teilnehmer - platz) * 10) / teilnehmer + 1);
}

export function sportabzeichenPunkte(platz: number): number {
	if (platz === 1) return 6;
	if (platz > 10) return 0.5;
	return (12 - platz) / 2;
}

/** Sortierfolge und Platzierung einer Klasse. */
export function klassenWertung(
	starter: readonly Starter[],
	regeln: Pick<Regeln, 'strafe1' | 'strafe2'>,
	veranstaltungsJahr: number | null = null
): WertungsZeile[] {
	const zeilen: WertungsZeile[] = starter.map((s) => {
		const ergebnisse = {
			0: laufErgebnis(s.laeufe[0], regeln),
			1: laufErgebnis(s.laeufe[1], regeln),
			2: laufErgebnis(s.laeufe[2], regeln)
		} as Record<LaufNr, number | null>;
		const vollstaendig = ergebnisse[1] !== null && ergebnisse[2] !== null;
		const gesamt = vollstaendig ? runde(ergebnisse[1]! + ergebnisse[2]!) : null;
		const bester = vollstaendig
			? Math.min(ergebnisse[1]!, ergebnisse[2]!)
			: (ergebnisse[1] ?? ergebnisse[2] ?? null);
		const status: WertungsStatus = s.ausserWertung ? 'ausser-wertung' : vollstaendig ? 'gewertet' : 'unvollstaendig';
		return {
			starter: s,
			ergebnisse,
			gesamt,
			bester,
			platz: null,
			status,
			punkte: 0,
			sportabzeichen: 0,
			rookie: istRookie(s, veranstaltungsJahr)
		};
	});

	const rang: Record<WertungsStatus, number> = { gewertet: 0, unvollstaendig: 1, 'ausser-wertung': 2 };
	zeilen.sort((a, b) => {
		if (a.status !== b.status) return rang[a.status] - rang[b.status];
		if (a.status === 'gewertet') {
			const d = hundertstel(a.gesamt!) - hundertstel(b.gesamt!) || hundertstel(a.bester!) - hundertstel(b.bester!);
			if (d !== 0) return d;
		}
		return a.starter.startnummer - b.starter.startnummer;
	});

	// Wie in Excel zählen alle gemeldeten Fahrer der Klasse als Teilnehmer.
	const teilnehmer = zeilen.length;
	let vorher: WertungsZeile | null = null;
	zeilen.forEach((z, index) => {
		if (z.status !== 'gewertet') return;
		const gleich =
			vorher !== null &&
			hundertstel(vorher.gesamt!) === hundertstel(z.gesamt!) &&
			hundertstel(vorher.bester!) === hundertstel(z.bester!);
		z.platz = gleich ? vorher!.platz : index + 1;
		z.punkte = wertungsPunkte(z.platz!, teilnehmer);
		z.sportabzeichen = sportabzeichenPunkte(z.platz!);
		vorher = z;
	});

	return zeilen;
}

/** Nächster offener Lauf eines Starters (für die Erfassung). */
export function offeneLaeufe(s: Starter): LaufNr[] {
	return ([0, 1, 2] as LaufNr[]).filter((nr) => s.laeufe[nr]?.zeit === null || s.laeufe[nr]?.zeit === undefined);
}
