import { describe, expect, it } from 'vitest';
import { mannschaftsWertung } from './mannschaft';
import { klassenWertung } from './wertung';
import type { Klasse } from './typen';
import { lauf, starter } from './testdaten';

const regeln = { strafe1: 2, strafe2: 10 };
const klasse = (id: number, inMannschaft = true): Klasse => ({ id, name: `Klasse ${id}`, kuerzel: `K${id}`, position: id, inMannschaft });

describe('Mannschaftswertung', () => {
	it('summiert die besten N Ergebnisse je Verein über alle Klassen', () => {
		const k1 = klassenWertung(
			[
				starter({ startnummer: 1, verein: 'MSC A', laeufe: { 1: lauf(30), 2: lauf(30) } }),
				starter({ startnummer: 2, verein: 'MSC B', laeufe: { 1: lauf(31), 2: lauf(31) } }),
				starter({ startnummer: 3, verein: 'msc  a ', laeufe: { 1: lauf(32), 2: lauf(32) } })
			],
			regeln
		);
		const k6 = klassenWertung([starter({ startnummer: 60, verein: 'MSC B', laeufe: { 1: lauf(30), 2: lauf(30) } })], regeln);
		const w = mannschaftsWertung(
			[
				{ klasse: klasse(1), zeilen: k1 },
				{ klasse: klasse(6), zeilen: k6 }
			],
			2
		);
		// K1: Platz 1 = 7,67 · Platz 2 = 4,33 · Platz 3 = 1; K6: Platz 1 = 1
		expect(w.map((z) => [z.verein, z.summe, z.platz])).toEqual([
			['MSC A', 8.67, 1],
			['MSC B', 5.33, 2]
		]);
		expect(w[1].ergebnisse.map((e) => e.klasse)).toEqual(['K1', 'K6']);
	});

	it('ignoriert Klassen ohne Mannschaftswertung', () => {
		const k = klassenWertung([starter({ startnummer: 1, verein: 'MSC A', laeufe: { 1: lauf(30), 2: lauf(30) } })], regeln);
		expect(mannschaftsWertung([{ klasse: klasse(1, false), zeilen: k }], 6)).toEqual([]);
	});

	it('führt punktgleiche Fahrer auf dem letzten zählenden Platz gemeinsam auf', () => {
		const zeilen = ['A', 'B', 'C'].map((name, i) =>
			klassenWertung([starter({ startnummer: i, nachname: name, vorname: '', verein: 'MSC X', laeufe: { 1: lauf(30), 2: lauf(30) } })], regeln)
		);
		const w = mannschaftsWertung(zeilen.map((z, i) => ({ klasse: klasse(i + 1), zeilen: z })), 2);
		expect(w[0].ergebnisse).toHaveLength(2);
		expect(w[0].ergebnisse[1].fahrer).toEqual(['B', 'C']);
		expect(w[0].summe).toBe(2);
	});

	it('vergibt bei gleicher Summe den gleichen Platz', () => {
		const k = klassenWertung(
			[
				starter({ startnummer: 1, verein: 'MSC A', laeufe: { 1: lauf(30), 2: lauf(31) } }),
				starter({ startnummer: 2, verein: 'MSC B', laeufe: { 1: lauf(31), 2: lauf(30) } })
			],
			regeln
		);
		expect(mannschaftsWertung([{ klasse: klasse(1), zeilen: k }], 6).map((z) => z.platz)).toEqual([1, 1]);
	});
});
