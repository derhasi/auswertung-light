import { beforeEach, describe, expect, it } from 'vitest';
import { migrieren } from './migration';
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
		expect(await migrieren(treiber)).toBe(1);
		expect(await migrieren(treiber)).toBe(1);
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
		expect(liste.find((f) => f.lizenz === '2')?.verein).toBe('MSC Neu');
		await expect(r.fahrerSpeichern(fahrerDaten('1'))).rejects.toThrow('bereits vergeben');
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
		const basis = { klasseId: klassen[0].id, lizenz: '1', nachname: 'A', vorname: 'B', verein: 'V', plz: '', ort: '', rookieJahr: 2026, ausserWertung: false };
		const s = await r.startAnlegen(id, { ...basis, startnummer: 1 });
		await expect(r.startAnlegen(id, { ...basis, startnummer: 1 })).rejects.toThrow('Startnummer 1');

		await r.laufSpeichern(s.id, 1, { fehler1: 1, fehler2: 0, zeit: 30.12 });
		await r.laufSpeichern(s.id, 1, { fehler1: 2, fehler2: 1, zeit: 31.5, importId: '17' });
		await r.laufSpeichern(s.id, 0, { fehler1: 0, fehler2: 0, zeit: 40 });
		await r.laufSpeichern(s.id, 0, null);
		await r.startAktualisieren(s.id, { ausserWertung: true });

		const [geladen] = (await r.veranstaltungLaden(id))!.starter;
		expect(geladen.ausserWertung).toBe(true);
		expect(geladen.laeufe).toEqual({ 1: { fehler1: 2, fehler2: 1, zeit: 31.5, importId: '17' } });

		await expect(r.klasseLoeschen(klassen[0].id)).rejects.toThrow('noch Fahrer');
		expect((await r.fahrerStarts()).get('1')?.[0]).toMatchObject({ veranstaltung: 'Test', klasse: 'K1', startnummer: 1 });

		const exportDaten = await r.veranstaltungExportieren(id);
		const kopie = await r.veranstaltungImportieren(JSON.parse(JSON.stringify(exportDaten)));
		const kopieDaten = (await r.veranstaltungLaden(kopie))!;
		expect(kopieDaten.starter[0].laeufe[1]?.zeit).toBe(31.5);
		expect(kopieDaten.starter[0].klasseId).toBe(kopieDaten.klassen[0].id);

		await r.veranstaltungLoeschen(id);
		expect(await r.veranstaltungLaden(id)).toBeNull();
		expect((await r.veranstaltungsListe()).map((v) => v.id)).toEqual([kopie]);
	});
});
