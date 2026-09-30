/**
 * Plattformdienste: Dateien öffnen/speichern, beobachten und drucken.
 * In der Desktop-App über Tauri-Plugins, im Browser über Standard-Web-APIs.
 */
import { invoke, isTauri } from '@tauri-apps/api/core';

export const istDesktop = () => isTauri();

export interface Dateifilter {
	name: string;
	endungen: string[];
}

export interface GeoeffneteDatei {
	name: string;
	/** Nur in der Desktop-App verfügbar. */
	pfad: string | null;
	bytes: Uint8Array;
}

function dateiname(pfad: string): string {
	return pfad.split(/[\\/]/).pop() ?? pfad;
}

function browserDateiWaehlen(filter: Dateifilter[]): Promise<File | null> {
	return new Promise((resolve) => {
		const input = document.createElement('input');
		input.type = 'file';
		input.accept = filter.flatMap((f) => f.endungen.map((e) => `.${e}`)).join(',');
		input.addEventListener('change', () => resolve(input.files?.[0] ?? null));
		input.addEventListener('cancel', () => resolve(null));
		input.click();
	});
}

/** Pfad einer Datei auswählen (nur Desktop), z. B. für die Zeitmessung. */
export async function pfadWaehlen(titel: string, filter: Dateifilter[]): Promise<string | null> {
	if (!isTauri()) return null;
	const { open } = await import('@tauri-apps/plugin-dialog');
	const pfad = await open({ title: titel, multiple: false, directory: false, filters: filter.map((f) => ({ name: f.name, extensions: f.endungen })) });
	return typeof pfad === 'string' ? pfad : null;
}

export async function dateiLesen(pfad: string): Promise<Uint8Array> {
	const { readFile } = await import('@tauri-apps/plugin-fs');
	return readFile(pfad);
}

export async function dateiOeffnen(titel: string, filter: Dateifilter[]): Promise<GeoeffneteDatei | null> {
	if (isTauri()) {
		const pfad = await pfadWaehlen(titel, filter);
		if (!pfad) return null;
		return { name: dateiname(pfad), pfad, bytes: await dateiLesen(pfad) };
	}
	const datei = await browserDateiWaehlen(filter);
	if (!datei) return null;
	return { name: datei.name, pfad: null, bytes: new Uint8Array(await datei.arrayBuffer()) };
}

/** Speichert Daten über den „Speichern unter"-Dialog. Liefert den Pfad/Dateinamen oder `null` bei Abbruch. */
export async function dateiSpeichern(
	vorschlag: string,
	inhalt: Uint8Array | string,
	filter: Dateifilter[]
): Promise<string | null> {
	const bytes = typeof inhalt === 'string' ? new TextEncoder().encode(inhalt) : inhalt;
	if (isTauri()) {
		const { save } = await import('@tauri-apps/plugin-dialog');
		const { writeFile } = await import('@tauri-apps/plugin-fs');
		const pfad = await save({ title: 'Speichern unter', defaultPath: vorschlag, filters: filter.map((f) => ({ name: f.name, extensions: f.endungen })) });
		if (!pfad) return null;
		await writeFile(pfad, bytes);
		return pfad;
	}
	const url = URL.createObjectURL(new Blob([bytes as BlobPart]));
	const a = document.createElement('a');
	a.href = url;
	a.download = vorschlag;
	a.click();
	setTimeout(() => URL.revokeObjectURL(url), 1000);
	return vorschlag;
}

/** Beobachtet eine Datei und ruft `beiAenderung` nach Änderungen auf. Liefert eine Abmeldefunktion. */
export async function dateiBeobachten(pfad: string, beiAenderung: () => void): Promise<() => void> {
	if (!isTauri()) return () => {};
	const { watch } = await import('@tauri-apps/plugin-fs');
	return watch(pfad, () => beiAenderung(), { delayMs: 400 });
}

export async function drucken(): Promise<void> {
	if (isTauri()) {
		try {
			await invoke('drucken');
			return;
		} catch (e) {
			console.warn('Nativer Druck nicht verfügbar, verwende window.print()', e);
		}
	}
	window.print();
}

export async function bildAlsDataUrl(): Promise<string | null> {
	const datei = await dateiOeffnen('Logo auswählen', [{ name: 'Bilder', endungen: ['png', 'jpg', 'jpeg', 'gif', 'svg', 'webp'] }]);
	if (!datei) return null;
	const endung = datei.name.split('.').pop()?.toLowerCase() ?? 'png';
	const mime = endung === 'svg' ? 'image/svg+xml' : endung === 'jpg' ? 'image/jpeg' : `image/${endung}`;
	let binaer = '';
	for (let i = 0; i < datei.bytes.length; i += 0x8000) binaer += String.fromCharCode(...datei.bytes.subarray(i, i + 0x8000));
	return `data:${mime};base64,${btoa(binaer)}`;
}

export const CSV_FILTER: Dateifilter[] = [{ name: 'CSV-Datei', endungen: ['csv', 'txt'] }];
export const JSON_FILTER: Dateifilter[] = [{ name: 'Veranstaltung (JSON)', endungen: ['json'] }];
