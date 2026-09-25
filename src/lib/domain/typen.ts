/** Fachliche Typen der Auswertung (unabhängig von Datenbank und Oberfläche). */

/** 0 = Training, 1 = Wertungslauf 1, 2 = Wertungslauf 2 */
export type LaufNr = 0 | 1 | 2;
export const LAEUFE: readonly LaufNr[] = [0, 1, 2];
export const WERTUNGSLAEUFE: readonly LaufNr[] = [1, 2];

export const LAUF_NAMEN: Record<LaufNr, string> = {
	0: 'Training',
	1: 'Wertungslauf 1',
	2: 'Wertungslauf 2'
};

export const LAUF_KURZ: Record<LaufNr, string> = {
	0: 'Training',
	1: 'Lauf 1',
	2: 'Lauf 2'
};

/** Strafregeln einer Veranstaltung. */
export interface Regeln {
	/** Strafsekunden je Fehler der ersten Fehlerart (z. B. Pylone). */
	strafe1: number;
	/** Strafsekunden je Fehler der zweiten Fehlerart (z. B. Tor). */
	strafe2: number;
	/** Anzahl der Fahrerergebnisse, die pro Verein in die Mannschaftswertung eingehen. */
	mannschaftAnzahl: number;
}

export const STANDARD_REGELN: Regeln = { strafe1: 2, strafe2: 10, mannschaftAnzahl: 6 };

export interface LaufEingabe {
	fehler1: number;
	fehler2: number;
	/** Gefahrene Zeit in Sekunden; `null` = noch keine Zeit erfasst. */
	zeit: number | null;
	/** Kennung aus der Zeitmessung, falls die Zeit importiert wurde. */
	importId?: string | null;
}

export interface Klasse {
	id: number;
	name: string;
	kuerzel: string;
	position: number;
	/** Geht die Klasse in die Mannschaftswertung ein? */
	inMannschaft: boolean;
}

export interface Starter {
	id: number;
	klasseId: number;
	startnummer: number;
	lizenz: string;
	nachname: string;
	vorname: string;
	verein: string;
	plz: string;
	ort: string;
	rookieJahr: number | null;
	/** „niW" – startet außer Wertung. */
	ausserWertung: boolean;
	laeufe: Partial<Record<LaufNr, LaufEingabe>>;
}

export function anzeigeName(s: Pick<Starter, 'nachname' | 'vorname'>): string {
	return [s.nachname, s.vorname].filter(Boolean).join(', ');
}
