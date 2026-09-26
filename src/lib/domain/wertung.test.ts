import { describe, expect, it } from 'vitest';
import { klassenWertung, laufErgebnis, sportabzeichenPunkte, wertungsPunkte } from './wertung';
import { lauf, starter } from './testdaten';

const regeln = { strafe1: 2, strafe2: 10 };

describe('Laufergebnis', () => {
	it('addiert Strafsekunden', () => {
		expect(laufErgebnis(lauf(30.12, 2, 1), regeln)).toBe(44.12);
		expect(laufErgebnis(lauf(null, 1), regeln)).toBeNull();
		expect(laufErgebnis(undefined, regeln)).toBeNull();
	});
});

describe('Punkte', () => {
	it('berechnet Wertungspunkte wie die Excel-Formel', () => {
		// ROUND((N - Platz) * 10 / N + 1; 2)
		expect(wertungsPunkte(1, 12)).toBe(10.17);
		expect(wertungsPunkte(12, 12)).toBe(1);
		expect(wertungsPunkte(2, 3)).toBe(4.33);
	});

	it('berechnet Sportabzeichenpunkte', () => {
		expect(sportabzeichenPunkte(1)).toBe(6);
		expect(sportabzeichenPunkte(2)).toBe(5);
		expect(sportabzeichenPunkte(3)).toBe(4.5);
		expect(sportabzeichenPunkte(10)).toBe(1);
		expect(sportabzeichenPunkte(11)).toBe(0.5);
	});
});

describe('Klassenwertung', () => {
	it('sortiert nach Gesamtzeit und vergibt Plätze', () => {
		const a = starter({ startnummer: 1, laeufe: { 1: lauf(30), 2: lauf(31) } });
		const b = starter({ startnummer: 2, laeufe: { 1: lauf(29), 2: lauf(30, 1) } });
		const c = starter({ startnummer: 3, laeufe: { 1: lauf(40), 2: lauf(40), 0: lauf(1) } });
		const w = klassenWertung([c, a, b], regeln);
		expect(w.map((z) => z.starter.startnummer)).toEqual([2, 1, 3]);
		expect(w.map((z) => z.platz)).toEqual([1, 2, 3]);
		expect(w[0].gesamt).toBe(61);
		expect(w[0].punkte).toBe(wertungsPunkte(1, 3));
	});

	it('entscheidet Gleichstand über den besseren Lauf', () => {
		const a = starter({ startnummer: 1, laeufe: { 1: lauf(30), 2: lauf(30) } });
		const b = starter({ startnummer: 2, laeufe: { 1: lauf(29), 2: lauf(31) } });
		const w = klassenWertung([a, b], regeln);
		expect(w.map((z) => z.starter.startnummer)).toEqual([2, 1]);
		expect(w.map((z) => z.platz)).toEqual([1, 2]);
	});

	it('vergibt bei völligem Gleichstand den gleichen Platz', () => {
		const a = starter({ startnummer: 1, laeufe: { 1: lauf(30), 2: lauf(31) } });
		const b = starter({ startnummer: 2, laeufe: { 1: lauf(31), 2: lauf(30) } });
		const c = starter({ startnummer: 3, laeufe: { 1: lauf(40), 2: lauf(40) } });
		const w = klassenWertung([a, b, c], regeln);
		expect(w.map((z) => z.platz)).toEqual([1, 1, 3]);
		expect(w[0].punkte).toBe(w[1].punkte);
	});

	it('vergleicht Zeiten ohne Fließkommafehler', () => {
		const a = starter({ startnummer: 1, laeufe: { 1: lauf(0.1), 2: lauf(0.2) } });
		const b = starter({ startnummer: 2, laeufe: { 1: lauf(0.2), 2: lauf(0.1) } });
		expect(klassenWertung([a, b], regeln).map((z) => z.platz)).toEqual([1, 1]);
	});

	it('führt unvollständige und außer Wertung startende Fahrer ohne Platz', () => {
		const a = starter({ startnummer: 1, laeufe: { 1: lauf(30) } });
		const b = starter({ startnummer: 2, ausserWertung: true, laeufe: { 1: lauf(10), 2: lauf(10) } });
		const c = starter({ startnummer: 3, laeufe: { 1: lauf(40), 2: lauf(40) } });
		const w = klassenWertung([a, b, c], regeln);
		expect(w.map((z) => [z.starter.startnummer, z.status, z.platz])).toEqual([
			[3, 'gewertet', 1],
			[1, 'unvollstaendig', null],
			[2, 'ausser-wertung', null]
		]);
		// Alle gemeldeten Fahrer zählen als Teilnehmer (wie in Excel).
		expect(w[0].punkte).toBe(wertungsPunkte(1, 3));
		expect(w[2].punkte).toBe(0);
	});

	it('markiert Rookies des Veranstaltungsjahres', () => {
		const a = starter({ startnummer: 1, rookieJahr: 2026 });
		const b = starter({ startnummer: 2, rookieJahr: 2025 });
		const w = klassenWertung([a, b], regeln, 2026);
		expect(w.map((z) => z.rookie)).toEqual([true, false]);
	});

	it('führt DNS und DSQ ohne Platz und Punkte', () => {
		const dns = { fehler1: 0, fehler2: 0, zeit: null, status: 'dns' as const, kommentar: 'Motorschaden' };
		const dsq = { fehler1: 0, fehler2: 0, zeit: null, status: 'dsq' as const, kommentar: 'Frühstart' };
		const a = starter({ startnummer: 1, laeufe: { 1: lauf(30), 2: dsq } });
		const b = starter({ startnummer: 2, laeufe: { 1: dns, 2: lauf(30) } });
		const c = starter({ startnummer: 3, laeufe: { 1: lauf(40), 2: lauf(40) } });
		const d = starter({ startnummer: 4, laeufe: { 0: dns, 1: lauf(41), 2: lauf(41) } });
		const w = klassenWertung([a, b, c, d], regeln);
		expect(w.map((z) => [z.starter.startnummer, z.status, z.platz, z.punkte])).toEqual([
			[3, 'gewertet', 1, wertungsPunkte(1, 4)],
			[4, 'gewertet', 2, wertungsPunkte(2, 4)],
			[2, 'nicht-gestartet', null, 0],
			[1, 'disqualifiziert', null, 0]
		]);
		expect(w[3].ergebnisse[1]).toBe(30);
		expect(w[3].ergebnisse[2]).toBeNull();
	});
});
