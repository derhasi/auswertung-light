/**
 * Datenzugriff. Hinweis: Abfragen über mehrere Tabellen werden bewusst in
 * TypeScript zusammengeführt, weil tauri-plugin-sql Zeilen als Objekte liefert
 * und gleichnamige Spalten (z. B. „name") in Joins sonst verloren gingen.
 */
import { and, asc, eq, inArray, sql } from 'drizzle-orm';
import { drizzle, type SqliteRemoteDatabase } from 'drizzle-orm/sqlite-proxy';
import * as schema from './schema';
import { fahrer, klasse, lauf, start, veranstaltung } from './schema';
import type { SqlTreiber } from './treiber';
import type { FahrerDaten } from '$lib/domain/fahrer-import';
import type { Klasse, LaufEingabe, LaufNr, Starter } from '$lib/domain/typen';
import { STANDARD_ZEITQUELLE, type ZeitquelleEinstellung } from '$lib/domain/zeitquelle';

export type Db = SqliteRemoteDatabase<typeof schema>;

export function erstelleDb(treiber: SqlTreiber): Db {
	return drizzle(
		async (anweisung, params, methode) => {
			if (methode === 'run') {
				await treiber.ausfuehren(anweisung, params);
				return { rows: [] };
			}
			const zeilen = await treiber.abfragen(anweisung, params);
			return { rows: methode === 'get' ? (zeilen[0] as unknown[]) : zeilen };
		},
		{ schema }
	);
}

export interface Fahrer extends FahrerDaten {
	id: number;
	geaendertAm: string;
}

export interface Veranstaltung {
	id: number;
	name: string;
	datum: string;
	ort: string;
	ausrichter: string;
	zpId: string;
	strafe1: number;
	strafe2: number;
	fehler1Name: string;
	fehler2Name: string;
	mannschaftAnzahl: number;
	urkundenPlaetze: number;
	logoLinks: string;
	logoRechts: string;
	zeitquelle: ZeitquelleEinstellung;
	erstelltAm: string;
}

export type VeranstaltungsStammdaten = Omit<Veranstaltung, 'id' | 'erstelltAm'>;

export interface VeranstaltungsUebersicht {
	id: number;
	name: string;
	datum: string;
	ort: string;
	starter: number;
	klassen: number;
}

export interface VeranstaltungsDaten {
	veranstaltung: Veranstaltung;
	klassen: Klasse[];
	starter: Starter[];
}

export type StartDaten = Omit<Starter, 'id' | 'laeufe'>;

export interface FahrerStart {
	veranstaltungId: number;
	veranstaltung: string;
	datum: string;
	klasse: string;
	startnummer: number;
}

/** Austauschformat, um eine Veranstaltung zwischen Rechnern zu übertragen. */
export interface VeranstaltungsExport {
	format: 'auswertung-light-veranstaltung';
	version: 1;
	veranstaltung: VeranstaltungsStammdaten;
	klassen: (Omit<Klasse, 'id'> & { id: number })[];
	starter: Starter[];
}

export const STANDARD_KLASSEN: Omit<Klasse, 'id'>[] = [1, 2, 3, 4, 5, 6].map((n) => ({
	name: `Klasse ${n}`,
	kuerzel: `K${n}`,
	position: n,
	inMannschaft: true
}));

const jetzt = () => new Date().toISOString();

function zeitquelleLesen(json: string): ZeitquelleEinstellung {
	if (!json) return { ...STANDARD_ZEITQUELLE };
	try {
		return { ...STANDARD_ZEITQUELLE, ...JSON.parse(json) };
	} catch {
		return { ...STANDARD_ZEITQUELLE };
	}
}

function zuVeranstaltung(z: schema.VeranstaltungZeile): Veranstaltung {
	return { ...z, zeitquelle: zeitquelleLesen(z.zeitquelle) };
}

function zuKlasse(z: schema.KlasseZeile): Klasse {
	return { id: z.id, name: z.name, kuerzel: z.kuerzel, position: z.position, inMannschaft: Boolean(z.inMannschaft) };
}

function zuStarter(z: schema.StartZeile, laeufe: schema.LaufZeile[] = []): Starter {
	const s: Starter = {
		id: z.id,
		klasseId: z.klasseId,
		startnummer: z.startnummer,
		lizenz: z.lizenz,
		nachname: z.nachname,
		vorname: z.vorname,
		verein: z.verein,
		plz: z.plz,
		ort: z.ort,
		rookieJahr: z.rookieJahr,
		ausserWertung: Boolean(z.ausserWertung),
		laeufe: {}
	};
	for (const l of laeufe) {
		s.laeufe[l.nr as LaufNr] = { fehler1: l.fehler1, fehler2: l.fehler2, zeit: l.zeit, importId: l.importId };
	}
	return s;
}

function stammdatenZeile(d: Partial<VeranstaltungsStammdaten>) {
	const { zeitquelle, ...rest } = d;
	return zeitquelle === undefined ? rest : { ...rest, zeitquelle: JSON.stringify(zeitquelle) };
}

function eindeutigkeitsFehler(e: unknown, meldung: string): never {
	const text = `${e} ${(e as { cause?: unknown })?.cause ?? ''}`.toLowerCase();
	if (text.includes('unique')) throw new Error(meldung);
	throw e;
}

export class Repository {
	constructor(readonly db: Db) {}

	// ── Fahrerdatenbank ─────────────────────────────────────────────

	async fahrerListe(): Promise<Fahrer[]> {
		return this.db.select().from(fahrer).orderBy(asc(fahrer.nachname), asc(fahrer.vorname));
	}

	async fahrerSpeichern(daten: FahrerDaten & { id?: number }): Promise<number> {
		const { id, ...werte } = daten;
		try {
			if (id) {
				await this.db
					.update(fahrer)
					.set({ ...werte, geaendertAm: jetzt() })
					.where(eq(fahrer.id, id));
				return id;
			}
			const [neu] = await this.db
				.insert(fahrer)
				.values({ ...werte, geaendertAm: jetzt() })
				.returning({ id: fahrer.id });
			return neu.id;
		} catch (e) {
			eindeutigkeitsFehler(e, `Die Lizenz ${daten.lizenz} ist bereits vergeben.`);
		}
	}

	async fahrerLoeschen(ids: number[]): Promise<void> {
		if (ids.length) await this.db.delete(fahrer).where(inArray(fahrer.id, ids));
	}

	/** Gleicht die Fahrerdatenbank mit einer Importliste ab (neu / geändert / unverändert). */
	async fahrerImportieren(liste: FahrerDaten[]): Promise<{ neu: number; aktualisiert: number; unveraendert: number }> {
		const vorhanden = new Map((await this.fahrerListe()).map((f) => [f.lizenz, f]));
		const felder: (keyof FahrerDaten)[] = ['klasse', 'nachname', 'vorname', 'rookieJahr', 'plz', 'ort', 'verein', 'geburtsdatum', 'alteLizenz'];
		let neu = 0;
		let aktualisiert = 0;
		const schreiben: FahrerDaten[] = [];
		for (const f of liste) {
			const alt = vorhanden.get(f.lizenz);
			if (!alt) neu++;
			else if (felder.some((feld) => (alt[feld] ?? '') !== (f[feld] ?? ''))) aktualisiert++;
			else continue;
			schreiben.push(f);
		}
		const zeitpunkt = jetzt();
		for (let i = 0; i < schreiben.length; i += 50) {
			await this.db
				.insert(fahrer)
				.values(schreiben.slice(i, i + 50).map((f) => ({ ...f, geaendertAm: zeitpunkt })))
				.onConflictDoUpdate({
					target: fahrer.lizenz,
					set: Object.fromEntries(
						[...felder, 'geaendertAm' as const].map((feld) => [feld, sql.raw(`excluded.${fahrer[feld].name}`)])
					)
				});
		}
		return { neu, aktualisiert, unveraendert: liste.length - neu - aktualisiert };
	}

	/** In welchen Veranstaltungen ist welche Lizenz gestartet? */
	async fahrerStarts(): Promise<Map<string, FahrerStart[]>> {
		const [starts, veranstaltungen, klassen] = await Promise.all([
			this.db.select({ lizenz: start.lizenz, vid: start.veranstaltungId, kid: start.klasseId, nr: start.startnummer }).from(start),
			this.db.select().from(veranstaltung),
			this.db.select().from(klasse)
		]);
		const vMap = new Map(veranstaltungen.map((v) => [v.id, v]));
		const kMap = new Map(klassen.map((k) => [k.id, k]));
		const ergebnis = new Map<string, FahrerStart[]>();
		for (const s of starts) {
			if (!s.lizenz) continue;
			const v = vMap.get(s.vid);
			if (!v) continue;
			const liste = ergebnis.get(s.lizenz) ?? [];
			liste.push({ veranstaltungId: v.id, veranstaltung: v.name, datum: v.datum, klasse: kMap.get(s.kid)?.kuerzel ?? '', startnummer: s.nr });
			ergebnis.set(s.lizenz, liste);
		}
		for (const liste of ergebnis.values()) liste.sort((a, b) => b.datum.localeCompare(a.datum));
		return ergebnis;
	}

	// ── Veranstaltungen ─────────────────────────────────────────────

	async veranstaltungsListe(): Promise<VeranstaltungsUebersicht[]> {
		const [liste, starter, klassen] = await Promise.all([
			this.db.select().from(veranstaltung),
			this.db.select({ vid: start.veranstaltungId, anzahl: sql<number>`count(*)` }).from(start).groupBy(start.veranstaltungId),
			this.db.select({ vid: klasse.veranstaltungId, anzahl: sql<number>`count(*)` }).from(klasse).groupBy(klasse.veranstaltungId)
		]);
		const sMap = new Map(starter.map((s) => [s.vid, Number(s.anzahl)]));
		const kMap = new Map(klassen.map((k) => [k.vid, Number(k.anzahl)]));
		return liste
			.map((v) => ({ id: v.id, name: v.name, datum: v.datum, ort: v.ort, starter: sMap.get(v.id) ?? 0, klassen: kMap.get(v.id) ?? 0 }))
			.sort((a, b) => b.datum.localeCompare(a.datum) || b.id - a.id);
	}

	/**
	 * Legt eine Veranstaltung an. Einstellungen (Strafen, Klassen, Logos, Zeitquelle)
	 * werden von der Vorlage übernommen – standardmäßig von der zuletzt angelegten Veranstaltung.
	 */
	async veranstaltungAnlegen(
		grunddaten: Pick<Veranstaltung, 'name' | 'datum'> & Partial<VeranstaltungsStammdaten>,
		vorlageId?: number | null
	): Promise<number> {
		let vorlage: Veranstaltung | null = null;
		let klassen: Omit<Klasse, 'id'>[] = STANDARD_KLASSEN;
		if (vorlageId === undefined) {
			const [letzte] = await this.db.select({ id: veranstaltung.id }).from(veranstaltung).orderBy(sql`${veranstaltung.id} desc`).limit(1);
			vorlageId = letzte?.id ?? null;
		}
		if (vorlageId) {
			const daten = await this.veranstaltungLaden(vorlageId);
			if (daten) {
				vorlage = daten.veranstaltung;
				if (daten.klassen.length) klassen = daten.klassen;
			}
		}
		const uebernommen: Partial<VeranstaltungsStammdaten> = vorlage
			? {
					ausrichter: vorlage.ausrichter,
					strafe1: vorlage.strafe1,
					strafe2: vorlage.strafe2,
					fehler1Name: vorlage.fehler1Name,
					fehler2Name: vorlage.fehler2Name,
					mannschaftAnzahl: vorlage.mannschaftAnzahl,
					urkundenPlaetze: vorlage.urkundenPlaetze,
					logoLinks: vorlage.logoLinks,
					logoRechts: vorlage.logoRechts,
					zeitquelle: vorlage.zeitquelle
				}
			: {};
		const [neu] = await this.db
			.insert(veranstaltung)
			.values({ ...stammdatenZeile({ ...uebernommen, ...grunddaten }), name: grunddaten.name, datum: grunddaten.datum, erstelltAm: jetzt() })
			.returning({ id: veranstaltung.id });
		for (const k of klassen) {
			await this.db.insert(klasse).values({ veranstaltungId: neu.id, name: k.name, kuerzel: k.kuerzel, position: k.position, inMannschaft: k.inMannschaft });
		}
		return neu.id;
	}

	async veranstaltungLaden(id: number): Promise<VeranstaltungsDaten | null> {
		const [v] = await this.db.select().from(veranstaltung).where(eq(veranstaltung.id, id));
		if (!v) return null;
		const [klassen, starts] = await Promise.all([
			this.db.select().from(klasse).where(eq(klasse.veranstaltungId, id)).orderBy(asc(klasse.position), asc(klasse.id)),
			this.db.select().from(start).where(eq(start.veranstaltungId, id)).orderBy(asc(start.startnummer))
		]);
		const laeufe = starts.length
			? await this.db
					.select()
					.from(lauf)
					.where(inArray(lauf.startId, starts.map((s) => s.id)))
			: [];
		const nachStart = new Map<number, schema.LaufZeile[]>();
		for (const l of laeufe) nachStart.set(l.startId, [...(nachStart.get(l.startId) ?? []), l]);
		return {
			veranstaltung: zuVeranstaltung(v),
			klassen: klassen.map(zuKlasse),
			starter: starts.map((s) => zuStarter(s, nachStart.get(s.id)))
		};
	}

	async veranstaltungAktualisieren(id: number, daten: Partial<VeranstaltungsStammdaten>): Promise<void> {
		await this.db.update(veranstaltung).set(stammdatenZeile(daten)).where(eq(veranstaltung.id, id));
	}

	async veranstaltungLoeschen(id: number): Promise<void> {
		const starts = await this.db.select({ id: start.id }).from(start).where(eq(start.veranstaltungId, id));
		if (starts.length) await this.db.delete(lauf).where(inArray(lauf.startId, starts.map((s) => s.id)));
		await this.db.delete(start).where(eq(start.veranstaltungId, id));
		await this.db.delete(klasse).where(eq(klasse.veranstaltungId, id));
		await this.db.delete(veranstaltung).where(eq(veranstaltung.id, id));
	}

	async veranstaltungExportieren(id: number): Promise<VeranstaltungsExport> {
		const daten = await this.veranstaltungLaden(id);
		if (!daten) throw new Error('Veranstaltung nicht gefunden.');
		const { id: _id, erstelltAm: _erstellt, ...stammdaten } = daten.veranstaltung;
		return { format: 'auswertung-light-veranstaltung', version: 1, veranstaltung: stammdaten, klassen: daten.klassen, starter: daten.starter };
	}

	async veranstaltungImportieren(daten: VeranstaltungsExport): Promise<number> {
		if (daten?.format !== 'auswertung-light-veranstaltung') throw new Error('Die Datei ist kein Veranstaltungs-Export von Auswertung Light.');
		const [neu] = await this.db
			.insert(veranstaltung)
			.values({ ...stammdatenZeile(daten.veranstaltung), name: daten.veranstaltung.name, datum: daten.veranstaltung.datum, erstelltAm: jetzt() })
			.returning({ id: veranstaltung.id });
		const klassenIds = new Map<number, number>();
		for (const k of daten.klassen) {
			const [nk] = await this.db
				.insert(klasse)
				.values({ veranstaltungId: neu.id, name: k.name, kuerzel: k.kuerzel, position: k.position, inMannschaft: k.inMannschaft })
				.returning({ id: klasse.id });
			klassenIds.set(k.id, nk.id);
		}
		for (const s of daten.starter) {
			const klasseId = klassenIds.get(s.klasseId);
			if (!klasseId) continue;
			const { id: _id, laeufe, ...rest } = s;
			const neuerStart = await this.startAnlegen(neu.id, { ...rest, klasseId });
			for (const [nr, eingabe] of Object.entries(laeufe)) {
				if (eingabe) await this.laufSpeichern(neuerStart.id, Number(nr) as LaufNr, eingabe);
			}
		}
		return neu.id;
	}

	// ── Klassen ─────────────────────────────────────────────────────

	async klasseAnlegen(veranstaltungId: number, daten: Omit<Klasse, 'id'>): Promise<Klasse> {
		const [k] = await this.db
			.insert(klasse)
			.values({ ...daten, veranstaltungId })
			.returning();
		return zuKlasse(k);
	}

	async klasseAktualisieren(id: number, daten: Partial<Omit<Klasse, 'id'>>): Promise<void> {
		await this.db.update(klasse).set(daten).where(eq(klasse.id, id));
	}

	async klasseLoeschen(id: number): Promise<void> {
		const [belegt] = await this.db.select({ id: start.id }).from(start).where(eq(start.klasseId, id)).limit(1);
		if (belegt) throw new Error('In dieser Klasse sind noch Fahrer gemeldet.');
		await this.db.delete(klasse).where(eq(klasse.id, id));
	}

	// ── Nennungen und Läufe ─────────────────────────────────────────

	async startAnlegen(veranstaltungId: number, daten: StartDaten): Promise<Starter> {
		try {
			const [s] = await this.db
				.insert(start)
				.values({ ...daten, veranstaltungId })
				.returning();
			return zuStarter(s);
		} catch (e) {
			eindeutigkeitsFehler(e, `Die Startnummer ${daten.startnummer} ist bereits vergeben.`);
		}
	}

	async startAktualisieren(id: number, daten: Partial<StartDaten>): Promise<void> {
		try {
			await this.db.update(start).set(daten).where(eq(start.id, id));
		} catch (e) {
			eindeutigkeitsFehler(e, `Die Startnummer ${daten.startnummer} ist bereits vergeben.`);
		}
	}

	async startLoeschen(id: number): Promise<void> {
		await this.db.delete(lauf).where(eq(lauf.startId, id));
		await this.db.delete(start).where(eq(start.id, id));
	}

	/** Speichert einen Lauf; `null` löscht die Eingabe. */
	async laufSpeichern(startId: number, nr: LaufNr, eingabe: LaufEingabe | null): Promise<void> {
		if (!eingabe) {
			await this.db.delete(lauf).where(and(eq(lauf.startId, startId), eq(lauf.nr, nr)));
			return;
		}
		const werte = {
			fehler1: eingabe.fehler1 || 0,
			fehler2: eingabe.fehler2 || 0,
			zeit: eingabe.zeit,
			importId: eingabe.importId ?? null,
			geaendertAm: jetzt()
		};
		await this.db
			.insert(lauf)
			.values({ startId, nr, ...werte })
			.onConflictDoUpdate({ target: [lauf.startId, lauf.nr], set: werte });
	}
}
