import { describe, expect, it } from 'vitest';
import { naechsterStart, offeneStarts, startReihenfolge } from './reihenfolge';
import { lauf, starter } from './testdaten';

const kurz = (liste: { starter: { startnummer: number }; lauf: number }[]) =>
	liste.map((p) => `${p.starter.startnummer}${['T', 'W1', 'W2'][p.lauf]}`);

describe('Startreihenfolge', () => {
	it('bildet Zweierpaare mit Training und Wertung 1, danach Wertung 2 für alle', () => {
		const liste = [3, 1, 5, 2, 4].map((nr) => starter({ startnummer: nr }));
		expect(kurz(startReihenfolge(liste))).toEqual([
			'1T', '2T', '1W1', '2W1',
			'3T', '4T', '3W1', '4W1',
			'5T', '5W1',
			'1W2', '2W2', '3W2', '4W2', '5W2'
		]);
	});

	it('sortiert nach Klassenreihenfolge und Startnummer', () => {
		const a = starter({ startnummer: 1, klasseId: 20 });
		const b = starter({ startnummer: 2, klasseId: 10 });
		const reihenfolge = startReihenfolge([a, b], new Map([[10, 1], [20, 2]]));
		expect(kurz(reihenfolge).slice(0, 2)).toEqual(['2T', '1T']);
	});

	it('springt zum nächsten offenen Start und überspringt Erfasstes', () => {
		const s1 = starter({ startnummer: 1, laeufe: { 0: lauf(30) } });
		const s2 = starter({ startnummer: 2, laeufe: { 0: lauf(31), 1: { fehler1: 0, fehler2: 0, zeit: null, status: 'dns', kommentar: 'x' } } });
		const reihenfolge = startReihenfolge([s1, s2]);
		expect(kurz([naechsterStart(reihenfolge)!])).toEqual(['1W1']);
		expect(kurz([naechsterStart(reihenfolge, { starterId: s1.id, lauf: 1 })!])).toEqual(['1W2']);
		expect(kurz(offeneStarts(reihenfolge))).toEqual(['1W1', '1W2', '2W2']);
	});

	it('liefert null, wenn alles erfasst ist', () => {
		const s = starter({ startnummer: 1, laeufe: { 0: lauf(1), 1: lauf(1), 2: lauf(1) } });
		expect(naechsterStart(startReihenfolge([s]))).toBeNull();
	});
});
