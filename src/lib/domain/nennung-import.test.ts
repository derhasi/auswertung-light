import { describe, expect, it } from 'vitest';
import * as XLSX from 'xlsx';
import { fahrerUnterschiede, leseNennungsDatei, parseNennungsTabelle, type FahrerDaten } from './fahrer-import';
import { lizenzGueltig } from './typen';

const basis: FahrerDaten = {
	lizenz: 'AB-12/3_x',
	klasse: 'K1',
	nachname: 'Muster',
	vorname: 'Max',
	rookieJahr: null,
	plz: '12345',
	ort: 'Musterstadt',
	verein: 'MSC Test',
	geburtsdatum: '2010-01-01',
	alteLizenz: ''
};

describe('Lizenznummern', () => {
	it('erlaubt Buchstaben, Ziffern und - / _', () => {
		expect(lizenzGueltig('012345')).toBe(true);
		expect(lizenzGueltig('AB-12/3_x')).toBe(true);
		expect(lizenzGueltig('AB 12')).toBe(false);
		expect(lizenzGueltig('12.3')).toBe(false);
		expect(lizenzGueltig('')).toBe(false);
	});
});

describe('Nennungs-Import', () => {
	it('liest Startnummer, Klasse und Fahrer ohne Lizenz', () => {
		const r = parseNennungsTabelle([
			['Start-Nr.', 'ID', 'Klasse', 'Nachname', 'Vorname', 'Verein'],
			['7', 'AB-1', 'K2', 'Muster', 'Max', 'MSC'],
			['', '', 'K2', 'Gast', 'Gustav', 'MSC'],
			['8', '', '', '', '', '']
		]);
		expect(r.mitKlasse).toBe(true);
		expect(r.zeilen.map((z) => [z.startnummer, z.lizenz, z.nachname, z.klasse])).toEqual([
			[7, 'AB-1', 'Muster', 'K2'],
			[null, '', 'Gast', 'K2']
		]);
		expect(r.fehler).toHaveLength(1);
	});

	it('liest Excel-Dateien', async () => {
		const mappe = XLSX.utils.book_new();
		XLSX.utils.book_append_sheet(
			mappe,
			XLSX.utils.aoa_to_sheet([
				['Nr', 'Lizenz', 'Nachname', 'Vorname'],
				[3, 12345, 'Muster', 'Max']
			]),
			'Nennung'
		);
		const bytes = new Uint8Array(XLSX.write(mappe, { type: 'array', bookType: 'xlsx' }));
		const r = await leseNennungsDatei('nennung.xlsx', bytes);
		expect(r.zeilen[0]).toMatchObject({ startnummer: 3, lizenz: '12345', nachname: 'Muster' });
		expect(r.mitKlasse).toBe(false);
	});
});

describe('Abweichungen', () => {
	it('erkennt abweichende Felder, ignoriert Schreibweise und leere Importwerte', () => {
		expect(fahrerUnterschiede(basis, { ...basis, nachname: ' muster ', klasse: 'K3' })).toEqual([]);
		expect(fahrerUnterschiede(basis, { ...basis, verein: 'MSC Neu', ort: '' })).toEqual(['verein']);
		expect(fahrerUnterschiede(basis, { ...basis, vorname: 'Moritz', plz: '99999' })).toEqual(['vorname', 'plz']);
	});
});
