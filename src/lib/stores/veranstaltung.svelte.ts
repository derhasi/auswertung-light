/**
 * Zustand einer geöffneten Veranstaltung. Alle Änderungen werden sofort
 * in die Datenbank geschrieben; Wertungen werden live abgeleitet.
 */
import { repo, type Fahrer, type FahrerRef, type StartDaten, type VeranstaltungsDaten, type VeranstaltungsStammdaten } from '$lib/db';
import type { FahrerDaten } from '$lib/domain/fahrer-import';
import { mannschaftsWertung } from '$lib/domain/mannschaft';
import { zusammengefuehrt, type ImportPosten } from '$lib/domain/nennung-import';
import { startReihenfolge } from '$lib/domain/reihenfolge';
import type { Klasse, LaufEingabe, LaufNr, Starter } from '$lib/domain/typen';
import { klassenWertung } from '$lib/domain/wertung';

export class VeranstaltungsStore {
	daten: VeranstaltungsDaten = $state()!;

	constructor(daten: VeranstaltungsDaten) {
		this.daten = daten;
	}

	get id() {
		return this.daten.veranstaltung.id;
	}
	get v() {
		return this.daten.veranstaltung;
	}
	get klassen() {
		return this.daten.klassen;
	}
	get starter() {
		return this.daten.starter;
	}

	jahr = $derived(Number(this.daten.veranstaltung.datum.slice(0, 4)) || null);

	klassenWertungen = $derived(
		this.daten.klassen.map((klasse) => ({
			klasse,
			zeilen: klassenWertung(
				this.daten.starter.filter((s) => s.klasseId === klasse.id),
				this.daten.veranstaltung,
				this.jahr
			)
		}))
	);

	mannschaft = $derived(mannschaftsWertung(this.klassenWertungen, this.daten.veranstaltung.mannschaftAnzahl));

	nachStartnummer = $derived(new Map(this.daten.starter.map((s) => [s.startnummer, s])));

	klassenPosition = $derived(new Map(this.daten.klassen.map((k, i) => [k.id, i])));

	/** Startreihenfolge des Veranstaltungstags (Paare: T, W1 – danach alle W2). */
	reihenfolge = $derived(startReihenfolge(this.daten.starter, this.klassenPosition));

	klasseVon(s: Pick<Starter, 'klasseId'>): Klasse | undefined {
		return this.daten.klassen.find((k) => k.id === s.klasseId);
	}

	// ── Stammdaten & Klassen ──────────────────────────────────────

	async aktualisieren(daten: Partial<VeranstaltungsStammdaten>) {
		await (await repo()).veranstaltungAktualisieren(this.id, daten);
		Object.assign(this.daten.veranstaltung, daten);
	}

	async klasseAnlegen() {
		const nr = Math.max(0, ...this.klassen.map((k) => k.position)) + 1;
		const k = await (await repo()).klasseAnlegen(this.id, { name: `Klasse ${nr}`, kuerzel: `K${nr}`, position: nr, inMannschaft: true });
		this.daten.klassen.push(k);
	}

	async klasseAktualisieren(id: number, daten: Partial<Omit<Klasse, 'id'>>) {
		await (await repo()).klasseAktualisieren(id, daten);
		const k = this.klassen.find((k) => k.id === id);
		if (k) Object.assign(k, daten);
	}

	async klasseLoeschen(id: number) {
		await (await repo()).klasseLoeschen(id);
		this.daten.klassen = this.klassen.filter((k) => k.id !== id);
	}

	async klasseVerschieben(id: number, richtung: -1 | 1) {
		const liste = [...this.klassen];
		const i = liste.findIndex((k) => k.id === id);
		const j = i + richtung;
		if (i < 0 || j < 0 || j >= liste.length) return;
		[liste[i], liste[j]] = [liste[j], liste[i]];
		const r = await repo();
		for (const [index, k] of liste.entries()) {
			if (k.position !== index + 1) await r.klasseAktualisieren(k.id, { position: index + 1 });
			k.position = index + 1;
		}
		this.daten.klassen = liste;
	}

	// ── Nennungen ──────────────────────────────────────────────────

	/** Vorschlag: nächste Nummer nach der höchsten der Klasse, sonst nach der höchsten insgesamt. */
	naechsteStartnummer(klasseId: number): number {
		const inKlasse = this.starter.filter((s) => s.klasseId === klasseId).map((s) => s.startnummer);
		const basis = inKlasse.length ? Math.max(...inKlasse) : Math.max(0, ...this.starter.map((s) => s.startnummer));
		let nr = basis + 1;
		while (this.nachStartnummer.has(nr)) nr++;
		return nr;
	}

	async nennen(daten: StartDaten): Promise<Starter> {
		const s = await (await repo()).startAnlegen(this.id, daten);
		this.daten.starter.push(s);
		this.daten.starter.sort((a, b) => a.startnummer - b.startnummer);
		return s;
	}

	async startAktualisieren(id: number, daten: Partial<StartDaten>) {
		await (await repo()).startAktualisieren(id, daten);
		const s = this.starter.find((s) => s.id === id);
		if (s) Object.assign(s, daten);
		if (daten.startnummer !== undefined) this.daten.starter.sort((a, b) => a.startnummer - b.startnummer);
	}

	async startLoeschen(id: number) {
		await (await repo()).startLoeschen(id);
		this.daten.starter = this.starter.filter((s) => s.id !== id);
	}

	/** Speichert einen Lauf. Änderungen an bereits erfassten Läufen brauchen einen Kommentar. */
	async laufSpeichern(startId: number, nr: LaufNr, eingabe: LaufEingabe | null, aenderungsKommentar?: string) {
		// Der gespeicherte Stand ist normalisiert (z. B. keine Zeit bei DNS/DSQ).
		const gespeichert = await (await repo()).laufSpeichern(startId, nr, eingabe, aenderungsKommentar);
		const s = this.starter.find((s) => s.id === startId);
		if (!s) return;
		if (gespeichert) s.laeufe[nr] = gespeichert;
		else delete s.laeufe[nr];
	}

	async laufAenderungen(startId: number) {
		return (await repo()).laufAenderungen(startId);
	}

	/**
	 * Nennt einen Fahrer aus der Datenbank bzw. legt ihn dort an und verknüpft die
	 * Nennung mit der aktuellen Version seiner Daten.
	 */
	async nennenMitFahrer(
		fahrerDaten: FahrerDaten,
		ref: FahrerRef,
		nennung: { klasseId: number; startnummer: number; ausserWertung?: boolean }
	): Promise<Starter> {
		return this.nennen({
			klasseId: nennung.klasseId,
			startnummer: nennung.startnummer,
			ausserWertung: nennung.ausserWertung ?? false,
			lizenz: fahrerDaten.lizenz,
			nachname: fahrerDaten.nachname,
			vorname: fahrerDaten.vorname,
			verein: fahrerDaten.verein,
			plz: fahrerDaten.plz,
			ort: fahrerDaten.ort,
			rookieJahr: fahrerDaten.rookieJahr,
			fahrerId: ref.id,
			fahrerVersionId: ref.versionId
		});
	}

	/** Führt einen geplanten Nennungs-Import aus. */
	async nennungenImportieren(posten: ImportPosten<Fahrer>[]): Promise<{ genannt: number; neueFahrer: number; aktualisiert: number; fehler: string[] }> {
		const r = await repo();
		const ergebnis = { genannt: 0, neueFahrer: 0, aktualisiert: 0, fehler: [] as string[] };
		for (const p of posten) {
			if (!p.uebernehmen || p.klasseId === null || p.startnummer === null) continue;
			try {
				let daten: FahrerDaten;
				let ref: FahrerRef;
				if (p.art === 'neu' || (p.art === 'konflikt' && p.loesung === 'neuer-fahrer')) {
					daten = p.zeile;
					ref = await r.fahrerSpeichern(p.zeile, p.art === 'neu' ? 'Nennungs-Import' : 'Nennungs-Import (anderer Fahrer)');
					ergebnis.neueFahrer++;
				} else if (p.art === 'konflikt' && p.loesung === 'zusammenfuehren' && p.bestand) {
					daten = zusammengefuehrt(p);
					ref = await r.fahrerSpeichern({ ...daten, id: p.bestand.id }, 'Nennungs-Import (zusammengeführt)');
					ergebnis.aktualisiert++;
				} else if (p.bestand) {
					daten = p.bestand;
					ref = await r.fahrerSpeichern({ ...p.bestand, id: p.bestand.id });
				} else continue;
				await this.nennenMitFahrer(daten, ref, { klasseId: p.klasseId, startnummer: p.startnummer });
				ergebnis.genannt++;
			} catch (e) {
				ergebnis.fehler.push(`${p.zeile.nachname}, ${p.zeile.vorname}: ${e instanceof Error ? e.message : e}`);
			}
		}
		return ergebnis;
	}
}

const cache = new Map<number, Promise<VeranstaltungsStore>>();

export function ladeVeranstaltung(id: number): Promise<VeranstaltungsStore> {
	let eintrag = cache.get(id);
	if (!eintrag) {
		eintrag = repo()
			.then((r) => r.veranstaltungLaden(id))
			.then((daten) => {
				if (!daten) throw new Error('Die Veranstaltung wurde nicht gefunden.');
				return new VeranstaltungsStore(daten);
			});
		eintrag.catch(() => cache.delete(id));
		cache.set(id, eintrag);
	}
	return eintrag;
}

export function veranstaltungVergessen(id: number) {
	cache.delete(id);
}
