/**
 * Zahlen- und Zeitformate.
 *
 * Intern wird mit Sekunden als Dezimalzahl gerechnet. Da Fließkommazahlen
 * Rundungsfehler haben, runden alle Berechnungen auf Hundertstel
 * (wie die Excel-Formeln der alten Auswertung mit ROUND(...; 2)).
 */

/** Rundet kaufmännisch auf n Nachkommastellen (wie Excel ROUND). */
export function runde(wert: number, stellen = 2): number {
	const faktor = 10 ** stellen;
	return Math.sign(wert) * (Math.round(Math.abs(wert) * faktor + 1e-9) / faktor);
}

/** Wandelt Sekunden in ganzzahlige Hundertstel (für exakte Vergleiche). */
export function hundertstel(sekunden: number): number {
	return Math.round(sekunden * 100 + (sekunden >= 0 ? 1e-9 : -1e-9));
}

/**
 * Liest eine Dezimalzahl mit Komma oder Punkt ein.
 * Leere Eingaben ergeben `null`, ungültige `undefined`.
 */
export function parseZahl(eingabe: string | number | null | undefined): number | null | undefined {
	if (eingabe === null || eingabe === undefined) return null;
	if (typeof eingabe === 'number') return Number.isFinite(eingabe) ? eingabe : undefined;
	const text = eingabe.trim().replace(/\s/g, '');
	if (text === '') return null;
	// Tausenderpunkte mit Dezimalkomma (1.234,56) ebenso wie 1234.56 erlauben.
	let normalisiert = text;
	if (normalisiert.includes(',') && normalisiert.includes('.')) {
		normalisiert = normalisiert.replace(/\./g, '').replace(',', '.');
	} else {
		normalisiert = normalisiert.replace(',', '.');
	}
	if (!/^-?\d*\.?\d+$|^-?\d+\.$/.test(normalisiert)) return undefined;
	const zahl = Number(normalisiert);
	return Number.isFinite(zahl) ? zahl : undefined;
}

/**
 * Liest eine Zeit in Sekunden ein. Erlaubt sind:
 * `62,34` · `62.34` · `1:02,34` · `1:02.34` · `0:01:02,34`
 */
export function parseZeit(eingabe: string | number | null | undefined): number | null | undefined {
	if (eingabe === null || eingabe === undefined) return null;
	if (typeof eingabe === 'number') return Number.isFinite(eingabe) && eingabe >= 0 ? runde(eingabe) : undefined;
	const text = eingabe.trim();
	if (text === '') return null;
	const teile = text.split(':');
	if (teile.length > 3) return undefined;
	const sekundenTeil = parseZahl(teile[teile.length - 1]);
	if (sekundenTeil === null || sekundenTeil === undefined || sekundenTeil < 0) return undefined;
	let sekunden = sekundenTeil;
	let faktor = 60;
	for (let i = teile.length - 2; i >= 0; i--) {
		if (!/^\d+$/.test(teile[i].trim())) return undefined;
		if (teile.length > 1 && sekundenTeil >= 60) return undefined;
		sekunden += Number(teile[i]) * faktor;
		faktor *= 60;
	}
	return runde(sekunden);
}

const dezimalFormat = new Intl.NumberFormat('de-DE', {
	minimumFractionDigits: 2,
	maximumFractionDigits: 2
});

const punkteFormat = new Intl.NumberFormat('de-DE', {
	minimumFractionDigits: 0,
	maximumFractionDigits: 2
});

/** Zeit/Ergebnis in deutscher Schreibweise mit zwei Nachkommastellen. */
export function formatZeit(sekunden: number | null | undefined): string {
	if (sekunden === null || sekunden === undefined) return '';
	return dezimalFormat.format(sekunden);
}

/** Punkte in deutscher Schreibweise ohne überflüssige Nullen. */
export function formatPunkte(punkte: number | null | undefined): string {
	if (punkte === null || punkte === undefined) return '';
	return punkteFormat.format(punkte);
}

/** Zahl mit Dezimalpunkt für Exporte (z. B. ZP-Format). */
export function punktZahl(wert: number | null | undefined, stellen?: number): string {
	if (wert === null || wert === undefined) return '';
	return stellen === undefined ? String(runde(wert)) : runde(wert, stellen).toFixed(stellen);
}
