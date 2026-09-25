import type { SqlTreiber } from './treiber';

const dateien = import.meta.glob('./migrationen/*.sql', { query: '?raw', import: 'default', eager: true }) as Record<
	string,
	string
>;

/** Migrationen in Dateireihenfolge (0000_…, 0001_…). */
export const MIGRATIONEN = Object.keys(dateien)
	.sort()
	.map((name) => ({ name, sql: dateien[name] }));

/** Führt ausstehende Migrationen aus; der Stand wird in PRAGMA user_version gehalten. */
export async function migrieren(treiber: SqlTreiber): Promise<number> {
	const [[version]] = (await treiber.abfragen('PRAGMA user_version')) as [[number]];
	let aktuell = Number(version) || 0;
	for (let i = aktuell; i < MIGRATIONEN.length; i++) {
		const anweisungen = MIGRATIONEN[i].sql
			.split('--> statement-breakpoint')
			.map((s) => s.trim())
			.filter(Boolean);
		for (const anweisung of anweisungen) await treiber.ausfuehren(anweisung);
		aktuell = i + 1;
		await treiber.ausfuehren(`PRAGMA user_version = ${aktuell}`);
	}
	return aktuell;
}
