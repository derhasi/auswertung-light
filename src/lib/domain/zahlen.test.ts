import { describe, expect, it } from 'vitest';
import { formatZeit, parseZahl, parseZeit, runde } from './zahlen';

describe('Zahlen', () => {
	it('rundet wie Excel', () => {
		expect(runde(1.005)).toBe(1.01);
		expect(runde(2.675)).toBe(2.68);
		expect(runde(0.1 + 0.2)).toBe(0.3);
	});

	it('liest Dezimalzahlen mit Komma und Punkt', () => {
		expect(parseZahl('12,34')).toBe(12.34);
		expect(parseZahl('12.34')).toBe(12.34);
		expect(parseZahl(' 7 ')).toBe(7);
		expect(parseZahl('')).toBeNull();
		expect(parseZahl('abc')).toBeUndefined();
	});

	it('liest Zeiten', () => {
		expect(parseZeit('62,34')).toBe(62.34);
		expect(parseZeit('1:02,34')).toBe(62.34);
		expect(parseZeit('0:01:02.5')).toBe(62.5);
		expect(parseZeit('1:75')).toBeUndefined();
		expect(parseZeit('-3')).toBeUndefined();
		expect(parseZeit('')).toBeNull();
	});

	it('formatiert deutsch', () => {
		expect(formatZeit(62.3)).toBe('62,30');
		expect(formatZeit(null)).toBe('');
	});
});
