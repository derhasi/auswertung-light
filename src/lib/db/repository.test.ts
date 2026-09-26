import { beforeEach, describe, expect, it } from 'vitest';
import { MIGRATIONEN, migrieren } from './migration';
import { erstelleDb, Repository } from './repository';
import { sqlJsTreiber } from './treiber';
import type { FahrerDaten } from '$lib/domain/fahrer-import';

async function neuesRepo() {
	const treiber = await sqlJsTreiber();
	await migrieren(treiber);
	return new Repository(erstelleDb(treiber));
}

const fahrerDaten = (lizenz: string, teil: Partial<FahrerDaten> = {}): FahrerDaten => ({
	lizenz,
	klasse: 'K1',
	nachname: 'Muster',
	vorname: 'Max',
	rookieJahr: null,
	plz: '12345',
	ort: 'Musterstadt',
	verein: 'MSC Test',
	geburtsdatum: '2010-01-01',
	alteLizenz: '',
	...teil
});

describe('Repository', () => {
	let r: Repository;
	beforeEach(async () => {
		r = await neuesRepo();
	});

	it('migriert idempotent', async () => {
		const treiber = await sqlJsTreiber();
		expect(await migrieren(treiber)).toBe(3);
		expect(await migrieren(treiber)).toBe(3);
	});

	it('übernimmt vorhandene Daten beim Update auf Fahrerversionen', async () => {
		const treiber = await sqlJsTreiber();
		const [erste] = MIGRATIONEN;
		for (const a of erste.sql.split('--> statement-breakpoint')) await treiber.ausfuehren(a);
		await treiber.ausfuehren('PRAGMA user_version = 1');
		await treiber.ausfuehren(`INSERT INTO fahrer (lizenz, nachname, geaendert_am) VALUES ('A-1', 'Alt', '2025-01-01')`);
		await treiber.ausfuehren(`INSERT INTO veranstaltung (name, datum, erstellt_am) VALUES ('V', '2025-05-01', 'x')`);
		await treiber.ausfuehren(`INSERT INTO klasse (veranstaltung_id, name) VALUES (1, 'K1')`);
		await treiber.ausfuehren(`INSERT INTO start (veranstaltung_id, klasse_id, startnummer, lizenz, nachname) VALUES (1, 1, 1, 'A-1', 'Alt')`);
		await treiber.ausfuehren(`INSERT INTO lauf (start_id, nr, zeit, geaendert_am) VALUES (1, 1, 30.5, 'x')`);
		expect(await migrieren(treiber)).toBe(3);
		const repo = new Repository(erstelleDb(treiber));
		const [f] = await repo.fahrerListe();
		const versionen = await repo.fahrerVersionen(f.id);
		expect(versionen).toHaveLength(1);
		const [s] = (await repo.veranstaltungLaden(1))!.starter;
		expect(s).toMatchObject({ fahrerId: f.id, fahrerVersionId: versionen[0].id });
		expect(s.laeufe[1]).toMatchObject({ zeit: 30.5, status: 'ok' });
	});

	it('importiert und gleicht Fahrer ab', async () => {
		expect(await r.fahrerImportieren([fahrerDaten('1'), fahrerDaten('2')])).toEqual({ neu: 2, aktualisiert: 0, unveraendert: 0 });
		expect(await r.fahrerImportieren([fahrerDaten('1'), fahrerDaten('2', { verein: 'MSC Neu' }), fahrerDaten('3')])).toEqual({
			neu: 1,
			aktualisiert: 1,
			unveraendert: 1
		});
		const liste = await r.fahrerListe();
		expect(liste).toHaveLength(3);
		const f2 = liste.find((f) => f.lizenz === '2')!;
		expect(f2.verein).toBe('MSC Neu');
		expect((await r.fahrerVersionen(f2.id)).map((v) => [v.verein, v.anlass])).toEqual([
			['MSC Neu', 'ZP-Import'],
			['MSC Test', 'ZP-Import']
		]);
	});

	it('versioniert Fahrerdaten und hält Lizenzen eindeutig', async () => {
		const a = await r.fahrerSpeichern(fahrerDaten('AB-1/2_x'));
		const unveraendert = await r.fahrerSpeichern({ ...fahrerDaten('AB-1/2_x'), id: a.id });
		expect(unveraendert).toEqual(a);
		const geaendert = await r.fahrerSpeichern({ ...fahrerDaten('AB-1/2_x', { ort: 'Neustadt' }), id: a.id }, 'Nennungs-Import');
		expect(geaendert.id).toBe(a.id);
		expect(geaendert.versionId).not.toBe(a.versionId);
		expect((await r.fahrerVersionen(a.id)).map((v) => v.ort)).toEqual(['Neustadt', 'Musterstadt']);
		// Anderer Fahrer mit derselben Lizenz ist nicht erlaubt …
		await expect(r.fahrerSpeichern(fahrerDaten('AB-1/2_x', { nachname: 'Anders' }))).rejects.toThrow('bereits an einen anderen Fahrer vergeben');
		const b = await r.fahrerSpeichern(fahrerDaten('C-3', { nachname: 'Anders' }));
		await expect(r.fahrerSpeichern({ ...fahrerDaten('AB-1/2_x'), id: b.id })).rejects.toThrow('vergeben');
		// … Fahrer ohne Lizenz dürfen mehrfach vorkommen
		await r.fahrerSpeichern(fahrerDaten('', { nachname: 'Gast1' }));
		await r.fahrerSpeichern(fahrerDaten('', { nachname: 'Gast2' }));
		expect((await r.fahrerListe()).filter((f) => f.lizenz === '')).toHaveLength(2);
	});

	it('macht bei der Migration doppelte Lizenzen eindeutig', async () => {
		const treiber = await sqlJsTreiber();
		for (const m of MIGRATIONEN.slice(0, 2)) for (const a of m.sql.split('--> statement-breakpoint')) await treiber.ausfuehren(a);
		await treiber.ausfuehren('PRAGMA user_version = 2');
		await treiber.ausfuehren(`INSERT INTO fahrer (lizenz, nachname, geaendert_am) VALUES ('X1', 'A', 'x'), ('X1', 'B', 'x'), ('', 'C', 'x'), ('', 'D', 'x')`);
		expect(await migrieren(treiber)).toBe(3);
		const liste = await new Repository(erstelleDb(treiber)).fahrerListe();
		expect(liste.map((f) => [f.nachname, f.lizenz])).toEqual([
			['A', 'X1'],
			['B', 'X1-doppelt2'],
			['C', ''],
			['D', '']
		]);
	});

	it('legt Veranstaltungen mit Standardklassen an und übernimmt Einstellungen', async () => {
		const id = await r.veranstaltungAnlegen({ name: 'Lauf 1', datum: '2026-05-01' });
		let daten = (await r.veranstaltungLaden(id))!;
		expect(daten.klassen.map((k) => k.kuerzel)).toEqual(['K1', 'K2', 'K3', 'K4', 'K5', 'K6']);
		expect(daten.veranstaltung.strafe1).toBe(2);

		await r.veranstaltungAktualisieren(id, { strafe1: 3, zeitquelle: { ...daten.veranstaltung.zeitquelle, pfad: 'C:/zeit.csv' } });
		await r.klasseAnlegen(id, { name: 'Klasse 7', kuerzel: 'K7', position: 7, inMannschaft: false });

		const id2 = await r.veranstaltungAnlegen({ name: 'Lauf 2', datum: '2026-06-01' });
		daten = (await r.veranstaltungLaden(id2))!;
		expect(daten.veranstaltung.strafe1).toBe(3);
		expect(daten.veranstaltung.zeitquelle.pfad).toBe('C:/zeit.csv');
		expect(daten.klassen).toHaveLength(7);
		expect(daten.klassen[6].inMannschaft).toBe(false);

		const liste = await r.veranstaltungsListe();
		expect(liste.map((v) => v.name)).toEqual(['Lauf 2', 'Lauf 1']);
	});

	it('speichert Nennungen und Läufe', async () => {
		const id = await r.veranstaltungAnlegen({ name: 'Test', datum: '2026-05-01' });
		const { klassen } = (await r.veranstaltungLaden(id))!;
		const f = await r.fahrerSpeichern(fahrerDaten('1', { nachname: 'A', vorname: 'B', verein: 'V' }));
		const basis = { klasseId: klassen[0].id, lizenz: '1', nachname: 'A', vorname: 'B', verein: 'V', plz: '', ort: '', rookieJahr: 2026, ausserWertung: false, fahrerId: f.id, fahrerVersionId: f.versionId };
		const s = await r.startAnlegen(id, { ...basis, startnummer: 1 });
		await expect(r.startAnlegen(id, { ...basis, startnummer: 1 })).rejects.toThrow('Startnummer 1');

		await r.laufSpeichern(s.id, 1, { fehler1: 1, fehler2: 0, zeit: 30.12 });
		// Änderung eines erfassten Laufs nur mit Begründung
		await expect(r.laufSpeichern(s.id, 1, { fehler1: 2, fehler2: 1, zeit: 31.5, importId: '17' })).rejects.toThrow('begründen');
		await r.laufSpeichern(s.id, 1, { fehler1: 2, fehler2: 1, zeit: 31.5, importId: '17' }, 'Tor übersehen');
		// Unveränderte Wiederholung braucht keine Begründung
		await r.laufSpeichern(s.id, 1, { fehler1: 2, fehler2: 1, zeit: 31.5, importId: '17' });
		await r.laufSpeichern(s.id, 0, { fehler1: 0, fehler2: 0, zeit: 40 });
		await r.laufSpeichern(s.id, 0, null, 'falscher Fahrer');
		await expect(r.laufSpeichern(s.id, 2, { fehler1: 0, fehler2: 0, zeit: null, status: 'dsq' })).rejects.toThrow('Kommentar');
		await r.laufSpeichern(s.id, 2, { fehler1: 3, fehler2: 0, zeit: 20, status: 'dsq', kommentar: 'Frühstart' });
		await r.startAktualisieren(s.id, { ausserWertung: true });

		const [geladen] = (await r.veranstaltungLaden(id))!.starter;
		expect(geladen.ausserWertung).toBe(true);
		expect(geladen.fahrerVersionId).toBe(f.versionId);
		expect(geladen.laeufe).toEqual({
			1: { fehler1: 2, fehler2: 1, zeit: 31.5, importId: '17', status: 'ok', kommentar: null },
			2: { fehler1: 0, fehler2: 0, zeit: null, importId: null, status: 'dsq', kommentar: 'Frühstart' }
		});
		const aenderungen = await r.laufAenderungen(s.id);
		expect(aenderungen.map((a) => [a.nr, a.kommentar, a.vorher?.zeit, a.nachher?.zeit ?? null])).toEqual([
			[0, 'falscher Fahrer', 40, null],
			[1, 'Tor übersehen', 30.12, 31.5]
		]);

		await expect(r.klasseLoeschen(klassen[0].id)).rejects.toThrow('noch Fahrer');
		expect((await r.fahrerStarts()).get(f.id)?.[0]).toMatchObject({ veranstaltung: 'Test', klasse: 'K1', startnummer: 1, fahrerVersionId: f.versionId });

		const exportDaten = await r.veranstaltungExportieren(id);
		const kopie = await r.veranstaltungImportieren(JSON.parse(JSON.stringify(exportDaten)));
		const kopieDaten = (await r.veranstaltungLaden(kopie))!;
		expect(kopieDaten.starter[0].laeufe[1]?.zeit).toBe(31.5);
		expect(kopieDaten.starter[0].klasseId).toBe(kopieDaten.klassen[0].id);
		expect(kopieDaten.starter[0].laeufe[2]?.status).toBe('dsq');
		expect(kopieDaten.starter[0].fahrerId).toBeNull();

		await r.veranstaltungLoeschen(id);
		expect(await r.veranstaltungLaden(id)).toBeNull();
		expect((await r.veranstaltungsListe()).map((v) => v.id)).toEqual([kopie]);
	});
});
