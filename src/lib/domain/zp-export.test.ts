import { describe, expect, it } from 'vitest';
import { klassenWertung } from './wertung';
import { klassenZiffer, zpExportCsv, zpExportZeilen } from './zp-export';
import { lauf, starter } from './testdaten';

describe('ZP-Export', () => {
	const klasse = { id: 3, name: 'Klasse 3', kuerzel: 'K3', position: 3, inMannschaft: true };

	it('erzeugt die 21 Spalten des alten zp_output-Blatts', () => {
		const zeilen = klassenWertung(
			[
				starter({ startnummer: 7, lizenz: '012345', nachname: 'Muster', vorname: 'Max', verein: 'MSC Test', laeufe: { 0: lauf(33.333, 1), 1: lauf(30.5, 1), 2: lauf(31.25, 0, 1) } }),
				starter({ startnummer: 8, lizenz: '5', ausserWertung: true, laeufe: { 1: lauf(20), 2: lauf(20) } })
			],
			{ strafe1: 2, strafe2: 10 }
		);
		const out = zpExportZeilen('4711', [{ klasse, zeilen }]);
		expect(out[0]).toEqual([
			'4711', '012345', 'MSC Test', '7', '3', '1',
			'1', '0', '33.33',
			'1', '0', '30.5',
			'0', '1', '31.25',
			'73.75', '1', '6', '6',
			'Muster, Max', '\\N'
		]);
		expect(out[1][5]).toBe('0');
		expect(out[1][16]).toBe('niW');
		expect(out[1][17]).toBe('0');
		expect(out[0]).toHaveLength(21);
		expect(zpExportCsv('1', [{ klasse, zeilen }])).toContain('"Muster, Max"');
	});

	it('führt DNS/DSQ als nicht gewertet ohne Platz', () => {
		const zeilen = klassenWertung(
			[starter({ startnummer: 9, laeufe: { 1: lauf(30), 2: { fehler1: 0, fehler2: 0, zeit: null, status: 'dsq', kommentar: 'x' } } })],
			{ strafe1: 2, strafe2: 10 }
		);
		const [z] = zpExportZeilen('1', [{ klasse, zeilen }]);
		expect([z[5], z[14], z[16], z[17]]).toEqual(['0', '0', '', '0']);
	});

	it('ermittelt die Klassenziffer', () => {
		expect(klassenZiffer({ kuerzel: 'K12', name: '' })).toBe('12');
		expect(klassenZiffer({ kuerzel: '', name: 'Klasse 6' })).toBe('6');
	});
});
