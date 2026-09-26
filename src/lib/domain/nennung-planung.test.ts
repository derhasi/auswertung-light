import { describe, expect, it } from 'vitest';
import type { FahrerDaten, NennungsZeile } from './fahrer-import';
import { importPlanen, konfliktGeloest, lizenzKorrigieren, offeneFelder, zusammengefuehrt, type DbFahrer } from './nennung-import';
import { starter } from './testdaten';
import type { Klasse } from './typen';

const daten = (lizenz: string, teil: Partial<FahrerDaten> = {}): FahrerDaten => ({
	lizenz,
	klasse: 'K1',
	nachname: 'Muster',
	vorname: 'Max',
	rookieJahr: null,
	plz: '12345',
	ort: 'Musterstadt',
	verein: 'MSC Test',
	geburtsdatum: '',
	alteLizenz: '',
	...teil
});
const zeile = (lizenz: string, teil: Partial<NennungsZeile> = {}): NennungsZeile => ({ ...daten(lizenz), startnummer: null, ...teil });
const klassen: Klasse[] = [
	{ id: 1, name: 'Klasse 1', kuerzel: 'K1', position: 1, inMannschaft: true },
	{ id: 2, name: 'Klasse 2', kuerzel: 'K2', position: 2, inMannschaft: true }
];

describe('Nennungs-Import planen', () => {
	const db: DbFahrer[] = [
		{ id: 10, ...daten('A1') },
		{ id: 11, ...daten('B2', { nachname: 'Bauer' }) },
		{ id: 12, ...daten('G3', { nachname: 'Gemeldet' }) }
	];
	const vorhanden = [starter({ startnummer: 5, klasseId: 1, lizenz: 'G3', fahrerId: 12 })];

	it('erkennt neue, bekannte, abweichende und bereits gemeldete Fahrer', () => {
		const plan = importPlanen(
			[zeile('A1'), zeile('B2', { nachname: 'Bauer', verein: 'MSC Neu' }), zeile('N4', { nachname: 'Neu' }), zeile('G3', { nachname: 'Gemeldet' })],
			db,
			vorhanden,
			klassen,
			null
		);
		expect(plan.map((p) => [p.zeile.lizenz, p.art, p.treffer, p.bestand?.id ?? null, p.unterschiede, p.uebernehmen])).toEqual([
			['A1', 'bekannt', 'lizenz', 10, [], true],
			['B2', 'konflikt', 'lizenz', 11, ['verein'], true],
			['N4', 'neu', null, null, [], true],
			['G3', 'bereits-gemeldet', 'lizenz', 12, [], false]
		]);
	});

	it('löst Konflikte auch bei gleichem Namen mit anderer Lizenz aus', () => {
		const [p] = importPlanen([zeile('Z9', { nachname: 'Bauer', ort: 'Anderswo' })], db, [], klassen, 1);
		expect([p.art, p.treffer, p.bestand?.id, p.unterschiede]).toEqual(['konflikt', 'name', 11, ['lizenz', 'ort']]);
	});

	it('verlangt eine Entscheidung für jedes abweichende Feld', () => {
		const [p] = importPlanen([zeile('B2', { nachname: 'Bauer', verein: 'MSC Neu', ort: 'Neustadt' })], db, [], klassen, 1);
		expect(konfliktGeloest(p)).toBe(false);
		expect(offeneFelder(p)).toEqual(['verein', 'ort']);
		expect(konfliktGeloest({ ...p, auswahl: { verein: 'import' } })).toBe(false);
		expect(konfliktGeloest({ ...p, auswahl: { verein: 'import', ort: 'bestand' } })).toBe(true);
		// Bei Lizenz-Treffer ist „anderer Fahrer“ nicht möglich (Lizenzen sind eindeutig)
		expect(konfliktGeloest({ ...p, loesung: 'neuer-fahrer' })).toBe(false);
		const [n] = importPlanen([zeile('Z9', { nachname: 'Bauer' })], db, [], klassen, 1);
		expect(konfliktGeloest({ ...n, loesung: 'neuer-fahrer' })).toBe(true);
	});

	it('vergibt Startnummern: gewünschte wenn frei, sonst nach der höchsten der Klasse', () => {
		const plan = importPlanen(
			[zeile('X1', { startnummer: 20 }), zeile('X2', { startnummer: 5 }), zeile('X3'), zeile('X4', { klasse: 'K2' })],
			[],
			vorhanden,
			klassen,
			null
		);
		expect(plan.map((p) => [p.startnummer, p.klasseId])).toEqual([
			[20, 1],
			[21, 1],
			[22, 1],
			[23, 2]
		]);
		expect(plan[1].nummerHinweis).toContain('vergeben');
	});

	it('nutzt die gewählte Zielklasse und meldet unbekannte Klassen', () => {
		expect(importPlanen([zeile('X', { klasse: 'K9' })], [], [], klassen, 2)[0].klasseId).toBe(2);
		const [p] = importPlanen([zeile('X', { klasse: 'K9' })], [], [], klassen, null);
		expect(p.uebernehmen).toBe(false);
		expect(p.hinweis).toContain('K9');
	});

	it('ordnet Fahrer ohne Lizenz über den Namen zu', () => {
		const ohne: DbFahrer[] = [{ id: 20, ...daten('', { nachname: 'Gast', vorname: 'Gustav' }) }];
		const [p] = importPlanen([zeile('', { nachname: 'gast', vorname: 'Gustav' })], ohne, [], klassen, 1);
		expect([p.art, p.treffer, p.bestand?.id]).toEqual(['bekannt', 'name', 20]);
	});

	it('führt Datensätze feldweise zusammen', () => {
		const [p] = importPlanen([zeile('B2', { nachname: 'Bauer', verein: 'MSC Neu', ort: 'Neustadt' })], db, [], klassen, 1);
		expect(p.unterschiede).toEqual(['verein', 'ort']);
		const m = zusammengefuehrt({ ...p, auswahl: { verein: 'import', ort: 'bestand' } });
		expect([m.verein, m.ort, m.lizenz]).toEqual(['MSC Neu', 'Musterstadt', 'B2']);
		expect('id' in m).toBe(false);
	});

	it('ignoriert leere Felder der Nennliste', () => {
		const [p] = importPlanen([zeile('A1', { verein: '', plz: '', ort: '' })], db, [], klassen, 1);
		expect([p.art, p.unterschiede]).toEqual(['bekannt', []]);
	});

	it('korrigiert eine versehentlich gleiche Lizenz und gleicht neu ab', () => {
		// Anderer Fahrer, aber per Tippfehler mit der Lizenz von Bauer (B2)
		const plan = importPlanen([zeile('B2', { nachname: 'Zeller', vorname: 'Zoe' }), zeile('Q1', { nachname: 'Quast' })], db, [], klassen, 1);
		expect([plan[0].art, plan[0].treffer]).toEqual(['konflikt', 'lizenz']);
		const neu = lizenzKorrigieren(plan, 0, 'B20', db, []);
		if (typeof neu === 'string') throw new Error(neu);
		expect([neu[0].art, neu[0].zeile.lizenz, neu[0].lizenzVorher, neu[0].uebernehmen]).toEqual(['neu', 'B20', 'B2', true]);
		// Korrektur auf die Lizenz eines anderen bekannten Fahrers → Abgleich mit diesem
		const zuA1 = lizenzKorrigieren(plan, 0, 'A1', db, []);
		if (typeof zuA1 === 'string') throw new Error(zuA1);
		expect([zuA1[0].art, zuA1[0].bestand?.id]).toEqual(['konflikt', 10]);
		// Ungültige oder doppelte Lizenz wird abgelehnt
		expect(lizenzKorrigieren(plan, 0, 'B 2', db, [])).toContain('Buchstaben');
		expect(lizenzKorrigieren(plan, 0, 'Q1', db, [])).toContain('bereits vor');
	});
});
