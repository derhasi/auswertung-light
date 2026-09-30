/**
 * Zeitimport aus der Datei einer Zeitmessanlage (CSV, XLS oder XLSX).
 * Entspricht dem alten Zeitimport, nur dass die Datei nicht mehr in Excel
 * geöffnet sein muss: Die App liest die Datei ein und beobachtet sie auf Änderungen.
 */
import { parseCsv, dekodiereText } from './csv';
import { parseZahl, parseZeit, runde } from './zahlen';

export type ZeitFormat = 'dezimal' | 'zeit';

export interface ZeitquelleEinstellung {
	/** Vollständiger Dateipfad (Desktop-App) bzw. Dateiname. */
	pfad: string;
	/** Tabellenblatt bei XLS/XLSX; leer = erstes Blatt. */
	blatt: string;
	/** Spalte mit der Kennung (1 = A); 0 = Zeilennummer als Kennung. */
	idSpalte: number;
	/** Spalte mit der Zeit (1 = A). */
	zeitSpalte: number;
	/** `dezimal` = Sekunden als Zahl; `zeit` = Uhrzeit/Dauer (Excel-Zeitwert oder „m:ss,cc"). */
	format: ZeitFormat;
	/** Anzahl der Kopfzeilen, die übersprungen werden. */
	kopfzeilen: number;
}

export const STANDARD_ZEITQUELLE: ZeitquelleEinstellung = {
	pfad: '',
	blatt: '',
	idSpalte: 1,
	zeitSpalte: 2,
	format: 'dezimal',
	kopfzeilen: 1
};

export interface GemesseneZeit {
	id: string;
	zeit: number;
	/** Zeilennummer in der Quelldatei (1-basiert). */
	zeile: number;
}

export type Zelle = string | number | boolean | Date | null | undefined;

export function zeitAusZelle(zelle: Zelle, format: ZeitFormat): number | null {
	if (zelle === null || zelle === undefined || zelle === '') return null;
	if (zelle instanceof Date) {
		const s = zelle.getHours() * 3600 + zelle.getMinutes() * 60 + zelle.getSeconds() + zelle.getMilliseconds() / 1000;
		return runde(s);
	}
	if (typeof zelle === 'boolean') return null;
	if (format === 'zeit') {
		if (typeof zelle === 'number') return runde(zelle * 86400);
		const z = parseZeit(zelle);
		return z ?? null;
	}
	const z = typeof zelle === 'number' ? zelle : parseZahl(zelle);
	return z === null || z === undefined || z < 0 ? null : runde(z);
}

export function zeitenAusTabelle(tabelle: readonly (readonly Zelle[])[], e: ZeitquelleEinstellung): GemesseneZeit[] {
	const zeiten: GemesseneZeit[] = [];
	tabelle.forEach((zeile, index) => {
		if (index < e.kopfzeilen) return;
		const zeit = zeitAusZelle(zeile[e.zeitSpalte - 1], e.format);
		if (zeit === null) return;
		const id = e.idSpalte === 0 ? String(index + 1) : String(zeile[e.idSpalte - 1] ?? '').trim();
		if (!id) return;
		zeiten.push({ id, zeit, zeile: index + 1 });
	});
	return zeiten;
}

export function istTabellenDatei(pfad: string): boolean {
	return /\.(xlsx|xlsm|xls|ods)$/i.test(pfad);
}

/** Liest die Zeiten aus Dateiinhalt (CSV direkt, Excel-Formate über SheetJS). */
export async function leseZeitquelle(bytes: Uint8Array, e: ZeitquelleEinstellung): Promise<GemesseneZeit[]> {
	if (!istTabellenDatei(e.pfad)) {
		return zeitenAusTabelle(parseCsv(dekodiereText(bytes)), e);
	}
	const XLSX = await import('xlsx');
	const mappe = XLSX.read(bytes, { type: 'array' });
	const blattName = e.blatt && mappe.SheetNames.includes(e.blatt) ? e.blatt : mappe.SheetNames[0];
	if (!blattName) return [];
	const tabelle = XLSX.utils.sheet_to_json<Zelle[]>(mappe.Sheets[blattName], { header: 1, raw: true, blankrows: true });
	return zeitenAusTabelle(tabelle, e);
}
