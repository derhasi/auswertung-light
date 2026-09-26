import { describe, expect, it } from 'vitest';
import type { FahrerDaten, NennungsZeile } from './fahrer-import';
import { importPlanen, zusammengefuehrt, type DbFahrer } from './nennung-import';
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
		expect(plan.map((p) => [p.zeile.lizenz, p.art, p.bestand?.id ?? null, p.unterschiede, p.uebernehmen])).toEqual([
			['A1', 'bekannt', 10, [], true],
			['B2', 'konflikt', 11, ['verein'], true],
			['N4', 'neu', null, [], true],
			['G3', 'bereits-gemeldet', 12, [], false]
		]);
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
		expect([p.art, p.bestand?.id]).toEqual(['bekannt', 20]);
	});

	it('führt Datensätze feldweise zusammen', () => {
		const [p] = importPlanen([zeile('B2', { nachname: 'Bauer', verein: 'MSC Neu', ort: 'Neustadt' })], db, [], klassen, 1);
		expect(p.unterschiede).toEqual(['verein', 'ort']);
		const m = zusammengefuehrt({ ...p, auswahl: { verein: 'import', ort: 'bestand' } });
		expect([m.verein, m.ort, m.lizenz]).toEqual(['MSC Neu', 'Musterstadt', 'B2']);
	});
});
