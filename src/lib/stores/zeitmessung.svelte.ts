/**
 * Anbindung der Zeitmessung: liest die konfigurierte Datei ein und
 * beobachtet sie in der Desktop-App auf Änderungen.
 */
import { leseZeitquelle, type GemesseneZeit, type ZeitquelleEinstellung } from '$lib/domain/zeitquelle';
import { dateiBeobachten, dateiLesen, dateiOeffnen, istDesktop } from '$lib/plattform';

export class Zeitmessung {
	zeiten = $state<GemesseneZeit[]>([]);
	stand = $state<Date | null>(null);
	fehler = $state<string | null>(null);
	private abmelden: (() => void) | null = null;
	private einstellung: ZeitquelleEinstellung;

	constructor(einstellung: ZeitquelleEinstellung) {
		this.einstellung = $state.snapshot(einstellung) as ZeitquelleEinstellung;
	}

	get konfiguriert() {
		return Boolean(this.einstellung.pfad);
	}

	get dateiname() {
		return this.einstellung.pfad.split(/[\\/]/).pop() ?? '';
	}

	/** Startet das Einlesen und – in der Desktop-App – die Dateibeobachtung. */
	async starten() {
		if (!this.konfiguriert || !istDesktop()) return;
		await this.einlesen();
		try {
			this.abmelden = await dateiBeobachten(this.einstellung.pfad, () => this.einlesen());
		} catch (e) {
			this.fehler = `Datei kann nicht beobachtet werden: ${e}`;
		}
	}

	async einlesen() {
		try {
			const bytes = await dateiLesen(this.einstellung.pfad);
			this.uebernehmen(await leseZeitquelle(bytes, this.einstellung));
		} catch (e) {
			this.fehler = `Zeitmessung konnte nicht gelesen werden: ${e}`;
		}
	}

	/** Browser-Modus: Datei einmalig auswählen und einlesen. */
	async dateiWaehlen() {
		const datei = await dateiOeffnen('Datei der Zeitmessung', [{ name: 'Zeitmessung', endungen: ['csv', 'txt', 'xlsx', 'xls', 'ods'] }]);
		if (!datei) return;
		this.einstellung = { ...this.einstellung, pfad: datei.pfad ?? datei.name };
		this.uebernehmen(await leseZeitquelle(datei.bytes, this.einstellung));
	}

	private uebernehmen(zeiten: GemesseneZeit[]) {
		this.zeiten = zeiten;
		this.stand = new Date();
		this.fehler = null;
	}

	beenden() {
		this.abmelden?.();
		this.abmelden = null;
	}
}
