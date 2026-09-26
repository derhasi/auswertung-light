/** Einfacher CSV-Leser/-Schreiber (RFC 4180) inkl. Zeichensatz-Erkennung. */

export type Trennzeichen = ',' | ';' | '\t';

/** Ermittelt das Trennzeichen anhand der ersten Zeile (außerhalb von Anführungszeichen). */
export function erkenneTrennzeichen(text: string): Trennzeichen {
	const zaehler: Record<Trennzeichen, number> = { ',': 0, ';': 0, '\t': 0 };
	let inQuotes = false;
	for (const zeichen of text) {
		if (zeichen === '"') inQuotes = !inQuotes;
		else if (!inQuotes && (zeichen === '\n' || zeichen === '\r')) break;
		else if (!inQuotes && zeichen in zaehler) zaehler[zeichen as Trennzeichen]++;
	}
	return (Object.entries(zaehler).sort((a, b) => b[1] - a[1])[0][0] as Trennzeichen) ?? ',';
}

export function parseCsv(text: string, trennzeichen: Trennzeichen = erkenneTrennzeichen(text)): string[][] {
	const zeilen: string[][] = [];
	let zeile: string[] = [];
	let feld = '';
	let inQuotes = false;
	text = text.replace(/^﻿/, '');

	for (let i = 0; i < text.length; i++) {
		const c = text[i];
		if (inQuotes) {
			if (c === '"') {
				if (text[i + 1] === '"') {
					feld += '"';
					i++;
				} else inQuotes = false;
			} else feld += c;
		} else if (c === '"') {
			inQuotes = true;
		} else if (c === trennzeichen) {
			zeile.push(feld);
			feld = '';
		} else if (c === '\n' || c === '\r') {
			if (c === '\r' && text[i + 1] === '\n') i++;
			zeile.push(feld);
			zeilen.push(zeile);
			zeile = [];
			feld = '';
		} else feld += c;
	}
	if (feld !== '' || zeile.length > 0) {
		zeile.push(feld);
		zeilen.push(zeile);
	}
	return zeilen.filter((z) => z.some((f) => f.trim() !== ''));
}

function feldCsv(wert: string | number | null | undefined, trennzeichen: string): string {
	const text = wert === null || wert === undefined ? '' : String(wert);
	return /["\r\n]/.test(text) || text.includes(trennzeichen) ? `"${text.replace(/"/g, '""')}"` : text;
}

export function stringifyCsv(zeilen: readonly (readonly (string | number | null | undefined)[])[], trennzeichen = ','): string {
	return zeilen.map((z) => z.map((f) => feldCsv(f, trennzeichen)).join(trennzeichen)).join('\r\n') + '\r\n';
}

/** Dekodiert eine Datei als UTF-8, bei ungültigen Zeichen als Windows-1252 (Excel-„ANSI"). */
export function dekodiereText(bytes: Uint8Array): string {
	try {
		return new TextDecoder('utf-8', { fatal: true }).decode(bytes).replace(/^﻿/, '');
	} catch {
		return new TextDecoder('windows-1252').decode(bytes);
	}
}

const CP1252_SONDERZEICHEN: Record<string, number> = {
	'€': 0x80, '‚': 0x82, 'ƒ': 0x83, '„': 0x84, '…': 0x85, '†': 0x86, '‡': 0x87, 'ˆ': 0x88,
	'‰': 0x89, 'Š': 0x8a, '‹': 0x8b, 'Œ': 0x8c, 'Ž': 0x8e, '‘': 0x91, '’': 0x92, '“': 0x93,
	'”': 0x94, '•': 0x95, '–': 0x96, '—': 0x97, '˜': 0x98, '™': 0x99, 'š': 0x9a, '›': 0x9b,
	'œ': 0x9c, 'ž': 0x9e, 'Ÿ': 0x9f
};

/** Kodiert Text als Windows-1252; nicht darstellbare Zeichen werden zu „?". */
export function kodiereWindows1252(text: string): Uint8Array {
	const bytes: number[] = [];
	for (const zeichen of text) {
		const code = zeichen.codePointAt(0)!;
		if (code < 0x80 || (code >= 0xa0 && code <= 0xff)) bytes.push(code);
		else bytes.push(CP1252_SONDERZEICHEN[zeichen] ?? 0x3f);
	}
	return Uint8Array.from(bytes);
}
