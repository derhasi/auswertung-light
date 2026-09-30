/**
 * Dünne Abstraktion über die SQLite-Anbindung:
 * - Desktop-App: tauri-plugin-sql (Datei im App-Datenverzeichnis)
 * - Browser (Entwicklung/Demo) und Tests: sql.js (SQLite als WebAssembly)
 */
import type { SqlJsStatic, Database as SqlJsDatabase } from 'sql.js';

export type SqlWert = string | number | null | Uint8Array;

export interface SqlTreiber {
	readonly art: 'tauri' | 'sqljs';
	ausfuehren(sql: string, params?: unknown[]): Promise<void>;
	/** Liefert die Zeilen als Arrays in Spaltenreihenfolge. */
	abfragen(sql: string, params?: unknown[]): Promise<unknown[][]>;
}

function wert(p: unknown): SqlWert {
	if (p === undefined || p === null) return null;
	if (typeof p === 'boolean') return p ? 1 : 0;
	if (typeof p === 'number' || typeof p === 'string' || p instanceof Uint8Array) return p;
	return String(p);
}

export async function tauriTreiber(pfad = 'sqlite:auswertung-light.db'): Promise<SqlTreiber> {
	const { default: Database } = await import('@tauri-apps/plugin-sql');
	const db = await Database.load(pfad);
	return {
		art: 'tauri',
		async ausfuehren(sql, params = []) {
			await db.execute(sql, params.map(wert));
		},
		async abfragen(sql, params = []) {
			const zeilen = await db.select<Record<string, unknown>[]>(sql, params.map(wert));
			return zeilen.map((z) => Object.values(z));
		}
	};
}

export interface SqlJsOptionen {
	/** Ort der sql-wasm.wasm-Datei (nur im Browser nötig). */
	wasmUrl?: string;
	/** Vorhandener Datenbankinhalt. */
	daten?: Uint8Array | null;
	/** Wird nach jeder Änderung mit dem Datenbankinhalt aufgerufen. */
	speichern?: (daten: Uint8Array) => void;
}

export async function sqlJsTreiber(optionen: SqlJsOptionen = {}): Promise<SqlTreiber & { db: SqlJsDatabase }> {
	const { default: initSqlJs } = (await import('sql.js')) as unknown as { default: (c?: object) => Promise<SqlJsStatic> };
	const SQL = await initSqlJs(optionen.wasmUrl ? { locateFile: () => optionen.wasmUrl } : undefined);
	const db = new SQL.Database(optionen.daten ?? undefined);
	db.run('PRAGMA foreign_keys = ON');

	// Sofort sichern: Ein verzögertes Speichern ginge beim Neuladen der Seite verloren.
	const merken = () => {
		if (!optionen.speichern) return;
		optionen.speichern(db.export());
		// export() öffnet die Datenbank neu und setzt dabei PRAGMAs zurück.
		db.run('PRAGMA foreign_keys = ON');
	};

	return {
		art: 'sqljs',
		db,
		async ausfuehren(sql, params = []) {
			db.run(sql, params.map(wert));
			merken();
		},
		async abfragen(sql, params = []) {
			const ergebnis = db.exec(sql, params.map(wert));
			if (!/^\s*select/i.test(sql)) merken();
			return ergebnis[0]?.values ?? [];
		}
	};
}
