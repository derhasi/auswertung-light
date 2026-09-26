/**
 * Datenzugriff. Hinweis: Abfragen über mehrere Tabellen werden bewusst in
 * TypeScript zusammengeführt, weil tauri-plugin-sql Zeilen als Objekte liefert
 * und gleichnamige Spalten (z. B. „name") in Joins sonst verloren gingen.
 */
import { and, asc, eq, inArray, sql } from 'drizzle-orm';
import { drizzle, type SqliteRemoteDatabase } from 'drizzle-orm/sqlite-proxy';
import * as schema from './schema';
import { fahrer, fahrerVersion, klasse, lauf, laufAenderung, start, veranstaltung } from './schema';
import type { SqlTreiber } from './treiber';
import type { FahrerDaten } from '$lib/domain/fahrer-import';
import type { Klasse, LaufEingabe, LaufNr, LaufStatus, Starter } from '$lib/domain/typen';
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

export interface FahrerVersion extends FahrerDaten {
	id: number;
	fahrerId: number;
	anlass: string;
	erstelltAm: string;
}

/** Verweis auf einen Fahrer und seine aktuelle Version. */
export interface FahrerRef {
	id: number;
	versionId: number;
}

export interface LaufAenderung {
	id: number;
	startId: number;
	nr: LaufNr;
	vorher: LaufEingabe | null;
	nachher: LaufEingabe | null;
	kommentar: string;
	zeitpunkt: string;
}

const FAHRER_FELDER: (keyof FahrerDaten)[] = ['lizenz', 'klasse', 'nachname', 'vorname', 'rookieJahr', 'plz', 'ort', 'verein', 'geburtsdatum', 'alteLizenz'];

function fahrerDaten(f: FahrerDaten): FahrerDaten {
	return Object.fromEntries(FAHRER_FELDER.map((feld) => [feld, f[feld] ?? (feld === 'rookieJahr' ? null : '')])) as unknown as FahrerDaten;
}

function fahrerGeaendert(a: FahrerDaten, b: FahrerDaten): boolean {
	return FAHRER_FELDER.some((feld) => (a[feld] ?? '') !== (b[feld] ?? ''));
}

function zuLaufEingabe(l: schema.LaufZeile): LaufEingabe {
	return {
		fehler1: l.fehler1,
		fehler2: l.fehler2,
		zeit: l.zeit,
		importId: l.importId,
		status: (l.status || 'ok') as LaufStatus,
		kommentar: l.kommentar
	};
}

/** Hat sich der Lauf inhaltlich geändert (Fehler, Zeit, Status, Kommentar)? */
function laufGeaendert(a: LaufEingabe | null, b: LaufEingabe | null): boolean {
	if (!a || !b) return a !== b;
	return (
		(a.fehler1 || 0) !== (b.fehler1 || 0) ||
		(a.fehler2 || 0) !== (b.fehler2 || 0) ||
		(a.zeit ?? null) !== (b.zeit ?? null) ||
		(a.status ?? 'ok') !== (b.status ?? 'ok') ||
		(a.kommentar ?? '') !== (b.kommentar ?? '')
	);
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
	fahrerVersionId: number | null;
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
		fahrerId: z.fahrerId,
		fahrerVersionId: z.fahrerVersionId,
		laeufe: {}
	};
	for (const l of laeufe) {
		s.laeufe[l.nr as LaufNr] = zuLaufEingabe(l);
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

	private async versionAnlegen(fahrerId: number, daten: FahrerDaten, anlass: string, zeitpunkt = jetzt()): Promise<number> {
		const [v] = await this.db
			.insert(fahrerVersion)
			.values({ ...fahrerDaten(daten), fahrerId, anlass, erstelltAm: zeitpunkt })
			.returning({ id: fahrerVersion.id });
		return v.id;
	}

	async aktuelleVersion(fahrerId: number): Promise<number | null> {
		const [v] = await this.db
			.select({ id: fahrerVersion.id })
			.from(fahrerVersion)
			.where(eq(fahrerVersion.fahrerId, fahrerId))
			.orderBy(sql`${fahrerVersion.id} desc`)
			.limit(1);
		return v?.id ?? null;
	}

	/**
	 * Legt einen Fahrer an oder aktualisiert ihn. Jede inhaltliche Änderung
	 * erzeugt eine neue Version; ältere Versionen bleiben erhalten.
	 */
	async fahrerSpeichern(daten: FahrerDaten & { id?: number }, anlass?: string): Promise<FahrerRef> {
		const { id, ...rest } = daten;
		const werte = fahrerDaten(rest as FahrerDaten);
		const zeitpunkt = jetzt();
		if (id) {
			const [alt] = await this.db.select().from(fahrer).where(eq(fahrer.id, id));
			if (!alt) throw new Error('Der Fahrer wurde nicht gefunden.');
			const versionId = await this.aktuelleVersion(id);
			if (!fahrerGeaendert(alt, werte) && versionId) return { id, versionId };
			try {
				await this.db
					.update(fahrer)
					.set({ ...werte, geaendertAm: zeitpunkt })
					.where(eq(fahrer.id, id));
			} catch (e) {
				eindeutigkeitsFehler(e, `Die Lizenz ${werte.lizenz} ist bereits an einen anderen Fahrer vergeben.`);
			}
			return { id, versionId: await this.versionAnlegen(id, werte, anlass ?? 'bearbeitet', zeitpunkt) };
		}
		let neu: { id: number };
		try {
			[neu] = await this.db
				.insert(fahrer)
				.values({ ...werte, geaendertAm: zeitpunkt })
				.returning({ id: fahrer.id });
		} catch (e) {
			eindeutigkeitsFehler(e, `Die Lizenz ${werte.lizenz} ist bereits an einen anderen Fahrer vergeben.`);
		}
		return { id: neu.id, versionId: await this.versionAnlegen(neu.id, werte, anlass ?? 'angelegt', zeitpunkt) };
	}

	async fahrerVersionen(fahrerId: number): Promise<FahrerVersion[]> {
		const zeilen = await this.db
			.select()
			.from(fahrerVersion)
			.where(eq(fahrerVersion.fahrerId, fahrerId))
			.orderBy(sql`${fahrerVersion.id} desc`);
		return zeilen;
	}

	/** Löscht Fahrer aus der Datenbank. Ihre Versionen bleiben für bestehende Nennungen erhalten. */
	async fahrerLoeschen(ids: number[]): Promise<void> {
		if (ids.length) await this.db.delete(fahrer).where(inArray(fahrer.id, ids));
	}

	/** Gleicht die Fahrerdatenbank über die (eindeutige) Lizenz mit der ZP-Fahrerliste ab. */
	async fahrerImportieren(liste: FahrerDaten[]): Promise<{ neu: number; aktualisiert: number; unveraendert: number }> {
		const nachLizenz = new Map((await this.fahrerListe()).filter((f) => f.lizenz).map((f) => [f.lizenz, f]));
		let aktualisiert = 0;
		let unveraendert = 0;
		const neue: FahrerDaten[] = [];
		for (const f of liste) {
			const alt = nachLizenz.get(f.lizenz);
			if (!alt) neue.push(fahrerDaten(f));
			else if (fahrerGeaendert(alt, fahrerDaten(f))) {
				await this.fahrerSpeichern({ ...f, id: alt.id }, 'ZP-Import');
				aktualisiert++;
			} else unveraendert++;
		}
		const zeitpunkt = jetzt();
		for (let i = 0; i < neue.length; i += 50) {
			const block = neue.slice(i, i + 50);
			const ids = await this.db
				.insert(fahrer)
				.values(block.map((f) => ({ ...f, geaendertAm: zeitpunkt })))
				.returning({ id: fahrer.id });
			await this.db
				.insert(fahrerVersion)
				.values(block.map((f, j) => ({ ...f, fahrerId: ids[j].id, anlass: 'ZP-Import', erstelltAm: zeitpunkt })));
		}
		return { neu: neue.length, aktualisiert, unveraendert };
	}

	/** In welchen Veranstaltungen ist welcher Fahrer (nach Fahrer-ID) gestartet? */
	async fahrerStarts(): Promise<Map<number, FahrerStart[]>> {
		const [starts, veranstaltungen, klassen] = await Promise.all([
			this.db
				.select({ fid: start.fahrerId, fvid: start.fahrerVersionId, vid: start.veranstaltungId, kid: start.klasseId, nr: start.startnummer })
				.from(start),
			this.db.select().from(veranstaltung),
			this.db.select().from(klasse)
		]);
		const vMap = new Map(veranstaltungen.map((v) => [v.id, v]));
		const kMap = new Map(klassen.map((k) => [k.id, k]));
		const ergebnis = new Map<number, FahrerStart[]>();
		for (const s of starts) {
			if (!s.fid) continue;
			const v = vMap.get(s.vid);
			if (!v) continue;
			const liste = ergebnis.get(s.fid) ?? [];
			liste.push({
				fahrerVersionId: s.fvid,
				veranstaltungId: v.id,
				veranstaltung: v.name,
				datum: v.datum,
				klasse: kMap.get(s.kid)?.kuerzel ?? '',
				startnummer: s.nr
			});
			ergebnis.set(s.fid, liste);
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
		if (starts.length) {
			await this.db.delete(laufAenderung).where(inArray(laufAenderung.startId, starts.map((s) => s.id)));
			await this.db.delete(lauf).where(inArray(lauf.startId, starts.map((s) => s.id)));
		}
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
			// Fahrer-IDs stammen aus einer anderen Datenbank und werden nicht übernommen.
			const neuerStart = await this.startAnlegen(neu.id, { ...rest, klasseId, fahrerId: null, fahrerVersionId: null });
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
		await this.db.delete(laufAenderung).where(eq(laufAenderung.startId, id));
		await this.db.delete(lauf).where(eq(lauf.startId, id));
		await this.db.delete(start).where(eq(start.id, id));
	}

	/**
	 * Speichert einen Lauf; `null` löscht die Eingabe.
	 * Wird ein bereits erfasster Lauf geändert oder gelöscht, ist ein Kommentar
	 * Pflicht – die Änderung wird mit altem und neuem Stand protokolliert.
	 */
	async laufSpeichern(startId: number, nr: LaufNr, eingabe: LaufEingabe | null, aenderungsKommentar?: string): Promise<LaufEingabe | null> {
		const [altZeile] = await this.db
			.select()
			.from(lauf)
			.where(and(eq(lauf.startId, startId), eq(lauf.nr, nr)));
		const vorher = altZeile ? zuLaufEingabe(altZeile) : null;
		const status = eingabe?.status ?? 'ok';
		const nachher: LaufEingabe | null = eingabe
			? {
					fehler1: status === 'ok' ? eingabe.fehler1 || 0 : 0,
					fehler2: status === 'ok' ? eingabe.fehler2 || 0 : 0,
					zeit: status === 'ok' ? eingabe.zeit : null,
					importId: status === 'ok' ? (eingabe.importId ?? null) : null,
					status,
					kommentar: eingabe.kommentar?.trim() || null
				}
			: null;
		if (nachher && status !== 'ok' && !nachher.kommentar) {
			throw new Error('Für DNS oder DSQ ist ein Kommentar erforderlich.');
		}
		const geaendert = vorher !== null && laufGeaendert(vorher, nachher);
		if (geaendert && !aenderungsKommentar?.trim()) {
			throw new Error('Bitte begründen Sie die Änderung des bereits erfassten Laufs.');
		}

		if (!nachher) {
			await this.db.delete(lauf).where(and(eq(lauf.startId, startId), eq(lauf.nr, nr)));
		} else {
			const werte = {
				fehler1: nachher.fehler1,
				fehler2: nachher.fehler2,
				zeit: nachher.zeit,
				importId: nachher.importId ?? null,
				status,
				kommentar: nachher.kommentar ?? null,
				geaendertAm: jetzt()
			};
			await this.db
				.insert(lauf)
				.values({ startId, nr, ...werte })
				.onConflictDoUpdate({ target: [lauf.startId, lauf.nr], set: werte });
		}
		if (geaendert) {
			await this.db.insert(laufAenderung).values({
				startId,
				nr,
				vorher: JSON.stringify(vorher),
				nachher: nachher ? JSON.stringify(nachher) : null,
				kommentar: aenderungsKommentar!.trim(),
				zeitpunkt: jetzt()
			});
		}
		return nachher;
	}

	async laufAenderungen(startId: number): Promise<LaufAenderung[]> {
		const zeilen = await this.db
			.select()
			.from(laufAenderung)
			.where(eq(laufAenderung.startId, startId))
			.orderBy(sql`${laufAenderung.id} desc`);
		return zeilen.map((z) => ({
			id: z.id,
			startId: z.startId,
			nr: z.nr as LaufNr,
			vorher: z.vorher ? (JSON.parse(z.vorher) as LaufEingabe) : null,
			nachher: z.nachher ? (JSON.parse(z.nachher) as LaufEingabe) : null,
			kommentar: z.kommentar,
			zeitpunkt: z.zeitpunkt
		}));
	}
}
