import { describe, expect, it } from 'vitest';
import * as XLSX from 'xlsx';
import { leseZeitquelle, zeitAusZelle, zeitenAusTabelle, STANDARD_ZEITQUELLE } from './zeitquelle';

describe('Zeitquelle', () => {
	it('liest Dezimal- und Zeitformate', () => {
		expect(zeitAusZelle('31,456', 'dezimal')).toBe(31.46);
		expect(zeitAusZelle(62.34 / 86400, 'zeit')).toBe(62.34);
		expect(zeitAusZelle('1:02,34', 'zeit')).toBe(62.34);
		expect(zeitAusZelle('', 'dezimal')).toBeNull();
	});

	it('ordnet Zeiten per ID oder Zeilennummer zu', () => {
		const tabelle = [['Nr', 'Zeit'], ['17', '30,10'], ['', ''], ['18', 'x']];
		expect(zeitenAusTabelle(tabelle, STANDARD_ZEITQUELLE)).toEqual([{ id: '17', zeit: 30.1, zeile: 2 }]);
		expect(zeitenAusTabelle(tabelle, { ...STANDARD_ZEITQUELLE, idSpalte: 0 })).toEqual([{ id: '2', zeit: 30.1, zeile: 2 }]);
	});

	it('liest CSV- und Excel-Dateien', async () => {
		const csv = new TextEncoder().encode('Nr;Zeit\n1;29,99\n2;30,01\n');
		expect(await leseZeitquelle(csv, { ...STANDARD_ZEITQUELLE, pfad: 'zeiten.csv' })).toHaveLength(2);

		const mappe = XLSX.utils.book_new();
		XLSX.utils.book_append_sheet(mappe, XLSX.utils.aoa_to_sheet([['Nr', 'Zeit'], [5, 45.678]]), 'Lauf');
		const bytes = new Uint8Array(XLSX.write(mappe, { type: 'array', bookType: 'xlsx' }));
		expect(await leseZeitquelle(bytes, { ...STANDARD_ZEITQUELLE, pfad: 'zeiten.xlsx', blatt: 'Lauf' })).toEqual([
			{ id: '5', zeit: 45.68, zeile: 2 }
		]);
	});
});
