/**
 * Zustand einer geöffneten Veranstaltung. Alle Änderungen werden sofort
 * in die Datenbank geschrieben; Wertungen werden live abgeleitet.
 */
import { repo, type StartDaten, type VeranstaltungsDaten, type VeranstaltungsStammdaten } from '$lib/db';
import { mannschaftsWertung } from '$lib/domain/mannschaft';
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

	async laufSpeichern(startId: number, nr: LaufNr, eingabe: LaufEingabe | null) {
		await (await repo()).laufSpeichern(startId, nr, eingabe);
		const s = this.starter.find((s) => s.id === startId);
		if (!s) return;
		if (eingabe) s.laeufe[nr] = { ...eingabe, importId: eingabe.importId ?? null };
		else delete s.laeufe[nr];
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
