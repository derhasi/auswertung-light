/** Globale Rückmeldungen (Toasts) und Bestätigungsdialoge. */

export type MeldungsArt = 'erfolg' | 'fehler' | 'info' | 'warnung';

export interface Meldung {
	id: number;
	art: MeldungsArt;
	text: string;
	details?: string[];
}

export interface Bestaetigung {
	titel: string;
	text: string;
	ja: string;
	nein: string;
	gefaehrlich: boolean;
	antwort: (ok: boolean) => void;
}

class UiZustand {
	meldungen = $state<Meldung[]>([]);
	bestaetigung = $state<Bestaetigung | null>(null);
	private naechste = 1;

	melden(text: string, art: MeldungsArt = 'erfolg', details?: string[]) {
		const id = this.naechste++;
		this.meldungen.push({ id, art, text, details });
		const dauer = art === 'fehler' || details?.length ? 9000 : 4000;
		setTimeout(() => this.schliessen(id), dauer);
	}

	fehler(e: unknown, prefix = '') {
		const text = e instanceof Error ? e.message : String(e);
		console.error(e);
		this.melden(prefix ? `${prefix}: ${text}` : text, 'fehler');
	}

	schliessen(id: number) {
		this.meldungen = this.meldungen.filter((m) => m.id !== id);
	}

	bestaetigen(
		text: string,
		optionen: { titel?: string; ja?: string; nein?: string; gefaehrlich?: boolean } = {}
	): Promise<boolean> {
		return new Promise((resolve) => {
			this.bestaetigung = {
				titel: optionen.titel ?? 'Bitte bestätigen',
				text,
				ja: optionen.ja ?? 'OK',
				nein: optionen.nein ?? 'Abbrechen',
				gefaehrlich: optionen.gefaehrlich ?? false,
				antwort: (ok) => {
					this.bestaetigung = null;
					resolve(ok);
				}
			};
		});
	}
}

export const ui = new UiZustand();
