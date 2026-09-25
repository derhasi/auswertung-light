import { isTauri } from '@tauri-apps/api/core';
import { migrieren } from './migration';
import { erstelleDb, Repository } from './repository';
import { sqlJsTreiber, tauriTreiber, type SqlTreiber } from './treiber';

export * from './repository';

const SPEICHER_SCHLUESSEL = 'auswertung-light-db';

let instanz: Promise<Repository> | null = null;

function ausSpeicher(): Uint8Array | null {
	try {
		const text = localStorage.getItem(SPEICHER_SCHLUESSEL);
		return text ? Uint8Array.from(atob(text), (c) => c.charCodeAt(0)) : null;
	} catch {
		return null;
	}
}

function inSpeicher(daten: Uint8Array) {
	try {
		let binaer = '';
		for (let i = 0; i < daten.length; i += 0x8000) binaer += String.fromCharCode(...daten.subarray(i, i + 0x8000));
		localStorage.setItem(SPEICHER_SCHLUESSEL, btoa(binaer));
	} catch (e) {
		console.warn('Datenbank konnte nicht im Browser gespeichert werden', e);
	}
}

async function oeffnen(): Promise<Repository> {
	let treiber: SqlTreiber;
	if (isTauri()) {
		treiber = await tauriTreiber();
	} else {
		// Browser-Modus für Entwicklung und Demo: Datenbank im localStorage.
		const { default: wasmUrl } = await import('sql.js/dist/sql-wasm.wasm?url');
		treiber = await sqlJsTreiber({ wasmUrl, daten: ausSpeicher(), speichern: inSpeicher });
	}
	await migrieren(treiber);
	return new Repository(erstelleDb(treiber));
}

/** Zentrale Datenbankinstanz der App. */
export function repo(): Promise<Repository> {
	instanz ??= oeffnen().catch((e) => {
		instanz = null;
		throw e;
	});
	return instanz;
}
