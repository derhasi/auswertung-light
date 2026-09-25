/**
 * Import der Fahrerdatenliste (CSV im Zugspitzpokal-Format).
 * Spalten: ID, Klasse, Nachname, Vorname, Rookie, PLZ, Wohnort, Verein, Geburtsdatum, Alte Lizenz-Nr.
 * Die Spalten werden anhand der Überschriften erkannt; fehlt eine erkennbare
 * Kopfzeile, gilt die obige Reihenfolge.
 */
import { parseCsv } from './csv';

export interface FahrerDaten {
	lizenz: string;
	klasse: string;
	nachname: string;
	vorname: string;
	rookieJahr: number | null;
	plz: string;
	ort: string;
	verein: string;
	/** ISO-Datum (JJJJ-MM-TT) oder Originaltext, falls nicht erkennbar. */
	geburtsdatum: string;
	alteLizenz: string;
}

type Feld = keyof FahrerDaten;

const REIHENFOLGE: Feld[] = ['lizenz', 'klasse', 'nachname', 'vorname', 'rookieJahr', 'plz', 'ort', 'verein', 'geburtsdatum', 'alteLizenz'];

const UEBERSCHRIFTEN: Record<string, Feld> = {
	id: 'lizenz',
	lizenz: 'lizenz',
	lizenznr: 'lizenz',
	lizenznummer: 'lizenz',
	klasse: 'klasse',
	nachname: 'nachname',
	name: 'nachname',
	vorname: 'vorname',
	rookie: 'rookieJahr',
	rookiejahr: 'rookieJahr',
	plz: 'plz',
	postleitzahl: 'plz',
	wohnort: 'ort',
	ort: 'ort',
	verein: 'verein',
	club: 'verein',
	geburtsdatum: 'geburtsdatum',
	geboren: 'geburtsdatum',
	altelizenznr: 'alteLizenz',
	altelizenz: 'alteLizenz'
};

function normalisiere(ueberschrift: string): string {
	return ueberschrift
		.toLowerCase()
		.replace(/[^a-zäöüß]/g, '');
}

/** Wandelt TT.MM.JJJJ in JJJJ-MM-TT; andere Formate bleiben unverändert. */
export function parseDatum(text: string): string {
	const t = text.trim();
	const de = /^(\d{1,2})\.(\d{1,2})\.(\d{4})$/.exec(t);
	if (de) return `${de[3]}-${de[2].padStart(2, '0')}-${de[1].padStart(2, '0')}`;
	return t;
}

export function formatDatum(iso: string | null | undefined): string {
	if (!iso) return '';
	const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(iso);
	return m ? `${m[3]}.${m[2]}.${m[1]}` : iso;
}

export interface ImportErgebnis {
	fahrer: FahrerDaten[];
	fehler: string[];
}

export function parseFahrerCsv(text: string): ImportErgebnis {
	const zeilen = parseCsv(text);
	const fehler: string[] = [];
	if (zeilen.length === 0) return { fahrer: [], fehler: ['Die Datei enthält keine Daten.'] };

	const kopf = zeilen[0].map((u) => UEBERSCHRIFTEN[normalisiere(u)]);
	const hatKopf = kopf.includes('lizenz') || kopf.filter(Boolean).length >= 3;
	const zuordnung: (Feld | undefined)[] = hatKopf ? kopf : REIHENFOLGE;
	const daten = hatKopf ? zeilen.slice(1) : zeilen;

	const gesehen = new Set<string>();
	const fahrer: FahrerDaten[] = [];
	daten.forEach((zeile, index) => {
		const zeilenNr = index + (hatKopf ? 2 : 1);
		const f: FahrerDaten = {
			lizenz: '', klasse: '', nachname: '', vorname: '', rookieJahr: null,
			plz: '', ort: '', verein: '', geburtsdatum: '', alteLizenz: ''
		};
		zuordnung.forEach((feld, spalte) => {
			if (!feld) return;
			const wert = (zeile[spalte] ?? '').trim();
			if (feld === 'rookieJahr') {
				const jahr = Number.parseInt(wert, 10);
				f.rookieJahr = Number.isFinite(jahr) && jahr > 1900 ? jahr : null;
			} else if (feld === 'geburtsdatum') {
				f.geburtsdatum = parseDatum(wert);
			} else {
				f[feld] = wert;
			}
		});
		if (!f.lizenz) {
			fehler.push(`Zeile ${zeilenNr}: keine Lizenznummer – übersprungen.`);
			return;
		}
		if (gesehen.has(f.lizenz)) {
			fehler.push(`Zeile ${zeilenNr}: Lizenz ${f.lizenz} doppelt – nur der erste Eintrag wird übernommen.`);
			return;
		}
		gesehen.add(f.lizenz);
		fahrer.push(f);
	});
	return { fahrer, fehler };
}

export const FAHRER_CSV_KOPF = ['ID', 'Klasse', 'Nachname', 'Vorname', 'Rookie', 'PLZ', 'Wohnort', 'Verein', 'Geburtsdatum', 'Alte Lizenz-Nr.'];

export function fahrerCsvZeile(f: FahrerDaten): string[] {
	return [f.lizenz, f.klasse, f.nachname, f.vorname, f.rookieJahr ? String(f.rookieJahr) : '', f.plz, f.ort, f.verein, formatDatum(f.geburtsdatum), f.alteLizenz];
}
