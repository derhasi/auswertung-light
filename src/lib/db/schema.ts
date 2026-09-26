import { integer, primaryKey, real, sqliteTable, text, uniqueIndex, index } from 'drizzle-orm/sqlite-core';

/**
 * Fahrerdatenbank – bleibt über alle Veranstaltungen hinweg erhalten.
 * Die Tabelle hält den aktuellen Stand; jede Änderung wird zusätzlich in
 * fahrer_version festgehalten. Lizenznummern sind nicht eindeutig: Zwei
 * verschiedene Fahrer können (z. B. durch Tippfehler) dieselbe Nummer tragen.
 */
export const fahrer = sqliteTable(
	'fahrer',
	{
		id: integer('id').primaryKey({ autoIncrement: true }),
		lizenz: text('lizenz').notNull(),
		klasse: text('klasse').notNull().default(''),
		nachname: text('nachname').notNull().default(''),
		vorname: text('vorname').notNull().default(''),
		rookieJahr: integer('rookie_jahr'),
		plz: text('plz').notNull().default(''),
		ort: text('ort').notNull().default(''),
		verein: text('verein').notNull().default(''),
		geburtsdatum: text('geburtsdatum').notNull().default(''),
		alteLizenz: text('alte_lizenz').notNull().default(''),
		geaendertAm: text('geaendert_am').notNull()
	},
	(t) => [index('fahrer_lizenz').on(t.lizenz)]
);

/** Historie der Fahrerdaten. Nennungen verweisen auf die Version, mit der gemeldet wurde. */
export const fahrerVersion = sqliteTable(
	'fahrer_version',
	{
		id: integer('id').primaryKey({ autoIncrement: true }),
		fahrerId: integer('fahrer_id').notNull(),
		lizenz: text('lizenz').notNull(),
		klasse: text('klasse').notNull().default(''),
		nachname: text('nachname').notNull().default(''),
		vorname: text('vorname').notNull().default(''),
		rookieJahr: integer('rookie_jahr'),
		plz: text('plz').notNull().default(''),
		ort: text('ort').notNull().default(''),
		verein: text('verein').notNull().default(''),
		geburtsdatum: text('geburtsdatum').notNull().default(''),
		alteLizenz: text('alte_lizenz').notNull().default(''),
		/** Woher die Version stammt (angelegt, bearbeitet, ZP-Import, Nennungs-Import …). */
		anlass: text('anlass').notNull().default(''),
		erstelltAm: text('erstellt_am').notNull()
	},
	(t) => [index('fahrer_version_fahrer').on(t.fahrerId)]
);

export const veranstaltung = sqliteTable('veranstaltung', {
	id: integer('id').primaryKey({ autoIncrement: true }),
	name: text('name').notNull(),
	/** ISO-Datum JJJJ-MM-TT */
	datum: text('datum').notNull(),
	ort: text('ort').notNull().default(''),
	ausrichter: text('ausrichter').notNull().default(''),
	/** Veranstaltungs-ID beim Zugspitzpokal-Ergebnisdienst (Spalte 1 im ZP-Export). */
	zpId: text('zp_id').notNull().default(''),
	strafe1: real('strafe1').notNull().default(2),
	strafe2: real('strafe2').notNull().default(10),
	fehler1Name: text('fehler1_name').notNull().default('Pylonen'),
	fehler2Name: text('fehler2_name').notNull().default('Tore'),
	mannschaftAnzahl: integer('mannschaft_anzahl').notNull().default(6),
	urkundenPlaetze: integer('urkunden_plaetze').notNull().default(3),
	/** Logos als Data-URL. */
	logoLinks: text('logo_links').notNull().default(''),
	logoRechts: text('logo_rechts').notNull().default(''),
	/** JSON mit den Einstellungen der Zeitquelle. */
	zeitquelle: text('zeitquelle').notNull().default(''),
	erstelltAm: text('erstellt_am').notNull()
});

export const klasse = sqliteTable(
	'klasse',
	{
		id: integer('id').primaryKey({ autoIncrement: true }),
		veranstaltungId: integer('veranstaltung_id')
			.notNull()
			.references(() => veranstaltung.id, { onDelete: 'cascade' }),
		name: text('name').notNull(),
		kuerzel: text('kuerzel').notNull().default(''),
		position: integer('position').notNull().default(0),
		inMannschaft: integer('in_mannschaft', { mode: 'boolean' }).notNull().default(true)
	},
	(t) => [index('klasse_veranstaltung').on(t.veranstaltungId)]
);

/** Nennung eines Fahrers. Die Fahrerdaten werden kopiert, damit spätere Änderungen alte Ergebnisse nicht verfälschen. */
export const start = sqliteTable(
	'start',
	{
		id: integer('id').primaryKey({ autoIncrement: true }),
		veranstaltungId: integer('veranstaltung_id')
			.notNull()
			.references(() => veranstaltung.id, { onDelete: 'cascade' }),
		klasseId: integer('klasse_id')
			.notNull()
			.references(() => klasse.id, { onDelete: 'cascade' }),
		startnummer: integer('startnummer').notNull(),
		lizenz: text('lizenz').notNull().default(''),
		nachname: text('nachname').notNull().default(''),
		vorname: text('vorname').notNull().default(''),
		verein: text('verein').notNull().default(''),
		plz: text('plz').notNull().default(''),
		ort: text('ort').notNull().default(''),
		rookieJahr: integer('rookie_jahr'),
		ausserWertung: integer('ausser_wertung', { mode: 'boolean' }).notNull().default(false),
		/** Verknüpfung zur Fahrerdatenbank (leer bei Fahrern ohne Datenbankeintrag). */
		fahrerId: integer('fahrer_id'),
		/** Version der Fahrerdaten, mit der gemeldet wurde. */
		fahrerVersionId: integer('fahrer_version_id')
	},
	(t) => [uniqueIndex('start_startnummer').on(t.veranstaltungId, t.startnummer), index('start_klasse').on(t.klasseId)]
);

export const lauf = sqliteTable(
	'lauf',
	{
		startId: integer('start_id')
			.notNull()
			.references(() => start.id, { onDelete: 'cascade' }),
		/** 0 = Training, 1/2 = Wertungsläufe */
		nr: integer('nr').notNull(),
		fehler1: integer('fehler1').notNull().default(0),
		fehler2: integer('fehler2').notNull().default(0),
		zeit: real('zeit'),
		importId: text('import_id'),
		/** ok, dns (nicht gestartet) oder dsq (disqualifiziert) */
		status: text('status').notNull().default('ok'),
		kommentar: text('kommentar'),
		geaendertAm: text('geaendert_am').notNull()
	},
	(t) => [primaryKey({ columns: [t.startId, t.nr] })]
);

/** Protokoll nachträglicher Änderungen an erfassten Läufen. */
export const laufAenderung = sqliteTable(
	'lauf_aenderung',
	{
		id: integer('id').primaryKey({ autoIncrement: true }),
		startId: integer('start_id')
			.notNull()
			.references(() => start.id, { onDelete: 'cascade' }),
		nr: integer('nr').notNull(),
		/** JSON des Laufs vor bzw. nach der Änderung (null = nicht vorhanden). */
		vorher: text('vorher'),
		nachher: text('nachher'),
		kommentar: text('kommentar').notNull(),
		zeitpunkt: text('zeitpunkt').notNull()
	},
	(t) => [index('lauf_aenderung_start').on(t.startId)]
);

export type FahrerZeile = typeof fahrer.$inferSelect;
export type VeranstaltungZeile = typeof veranstaltung.$inferSelect;
export type KlasseZeile = typeof klasse.$inferSelect;
export type StartZeile = typeof start.$inferSelect;
export type LaufZeile = typeof lauf.$inferSelect;
export type FahrerVersionZeile = typeof fahrerVersion.$inferSelect;
export type LaufAenderungZeile = typeof laufAenderung.$inferSelect;
