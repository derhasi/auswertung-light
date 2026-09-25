import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { dekodiereText, erkenneTrennzeichen, kodiereWindows1252, parseCsv, stringifyCsv } from './csv';
import { formatDatum, parseFahrerCsv } from './fahrer-import';

describe('CSV', () => {
	it('liest Anführungszeichen, Zeilenumbrüche und Trennzeichen', () => {
		expect(parseCsv('a,"b,c","d ""x"""\r\n1,2,3\n')).toEqual([
			['a', 'b,c', 'd "x"'],
			['1', '2', '3']
		]);
		expect(erkenneTrennzeichen('a;b;"c,d"\n')).toBe(';');
	});

	it('schreibt CSV und quotet nur bei Bedarf', () => {
		expect(stringifyCsv([['a', 'Muster, Max', 1.5]])).toBe('a,"Muster, Max",1.5\r\n');
	});

	it('erkennt Windows-1252 und kodiert wieder zurück', () => {
		const bytes = kodiereWindows1252('Dießen – Größe €');
		expect(dekodiereText(bytes)).toBe('Dießen – Größe €');
		expect(dekodiereText(new TextEncoder().encode('﻿München'))).toBe('München');
	});
});

describe('Fahrer-Import', () => {
	it('liest die Beispiel-CSV', () => {
		const text = readFileSync(new URL('../../../beispiele/fahrer-beispiel.csv', import.meta.url), 'utf8');
		const { fahrer, fehler } = parseFahrerCsv(text);
		expect(fehler).toEqual([]);
		expect(fahrer).toHaveLength(2);
		expect(fahrer[0]).toEqual({
			lizenz: '012345',
			klasse: 'K1',
			nachname: 'Mustermann',
			vorname: 'Martina',
			rookieJahr: 2016,
			plz: '12345',
			ort: 'Musterstadt',
			verein: 'MSC Musterhausen',
			geburtsdatum: '2009-01-01',
			alteLizenz: ''
		});
		expect(formatDatum(fahrer[0].geburtsdatum)).toBe('01.01.2009');
	});

	it('akzeptiert Dateien ohne Kopfzeile und meldet Fehler', () => {
		const { fahrer, fehler } = parseFahrerCsv('1;K2;Muster;Max;;;;MSC;\n;K1;Ohne;Lizenz\n1;K3;Doppelt;X\n');
		expect(fahrer).toHaveLength(1);
		expect(fahrer[0].klasse).toBe('K2');
		expect(fehler).toHaveLength(2);
	});
});
