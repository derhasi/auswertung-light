import type { LaufEingabe, Starter } from './typen';

let naechsteId = 1;

export function lauf(zeit: number | null, fehler1 = 0, fehler2 = 0): LaufEingabe {
	return { zeit, fehler1, fehler2 };
}

export function starter(teil: Partial<Starter> & { startnummer: number }): Starter {
	const id = naechsteId++;
	return {
		id,
		klasseId: 1,
		lizenz: String(1000 + id),
		nachname: `Fahrer${id}`,
		vorname: 'Test',
		verein: 'MSC Test',
		plz: '12345',
		ort: 'Musterstadt',
		rookieJahr: null,
		ausserWertung: false,
		laeufe: {},
		...teil
	};
}
