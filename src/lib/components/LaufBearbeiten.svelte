<script lang="ts">
	import { History, Save } from '@lucide/svelte';
	import type { LaufAenderung } from '$lib/db';
	import { anzeigeName, LAEUFE, LAUF_NAMEN, laufErfasst, type LaufEingabe, type LaufNr, type LaufStatus, type Starter } from '$lib/domain/typen';
	import { laufErgebnis } from '$lib/domain/wertung';
	import { formatZeit, parseZeit } from '$lib/domain/zahlen';
	import type { VeranstaltungsStore } from '$lib/stores/veranstaltung.svelte';
	import Dialog from '$lib/ui/Dialog.svelte';
	import { ui } from '$lib/ui/ui-zustand.svelte';

	interface Props {
		store: VeranstaltungsStore;
		starter: Starter | null;
		onschliessen: () => void;
	}

	let { store, starter, onschliessen }: Props = $props();

	interface Formular {
		status: LaufStatus;
		fehler1: number | null;
		fehler2: number | null;
		zeitText: string;
		kommentar: string;
	}

	let formular = $state<Record<LaufNr, Formular>>({} as Record<LaufNr, Formular>);
	let grund = $state('');
	let aenderungen = $state<LaufAenderung[]>([]);
	let speichernAktiv = $state(false);

	function ausLauf(l: LaufEingabe | undefined): Formular {
		return {
			status: l?.status ?? 'ok',
			fehler1: l?.fehler1 ?? 0,
			fehler2: l?.fehler2 ?? 0,
			zeitText: l?.zeit != null ? formatZeit(l.zeit) : '',
			kommentar: l?.kommentar ?? ''
		};
	}

	$effect(() => {
		if (!starter) return;
		formular = { 0: ausLauf(starter.laeufe[0]), 1: ausLauf(starter.laeufe[1]), 2: ausLauf(starter.laeufe[2]) };
		grund = '';
		store.laufAenderungen(starter.id).then((a) => (aenderungen = a)).catch((e) => ui.fehler(e));
	});

	/** Formular → Laufeingabe; `null` = kein Ergebnis (Lauf leer). */
	function eingabe(nr: LaufNr): LaufEingabe | null | undefined {
		const f = formular[nr];
		if (!f) return undefined;
		if (f.status !== 'ok') return { status: f.status, fehler1: 0, fehler2: 0, zeit: null, kommentar: f.kommentar.trim() || null };
		const zeit = parseZeit(f.zeitText);
		if (zeit === undefined) return undefined;
		if (zeit === null) return null;
		const alt = starter?.laeufe[nr];
		return {
			status: 'ok',
			fehler1: f.fehler1 ?? 0,
			fehler2: f.fehler2 ?? 0,
			zeit,
			importId: alt?.zeit === zeit ? (alt?.importId ?? null) : null,
			kommentar: f.kommentar.trim() || null
		};
	}

	function geaendert(nr: LaufNr): boolean {
		const alt = starter?.laeufe[nr];
		const neu = eingabe(nr);
		if (neu === undefined) return true;
		if (!laufErfasst(alt)) return neu !== null;
		if (neu === null) return true;
		return (
			(alt!.status ?? 'ok') !== neu.status ||
			(alt!.fehler1 || 0) !== neu.fehler1 ||
			(alt!.fehler2 || 0) !== neu.fehler2 ||
			(alt!.zeit ?? null) !== neu.zeit ||
			(alt!.kommentar ?? '') !== (neu.kommentar ?? '')
		);
	}

	const geaenderteLaeufe = $derived(starter && formular[0] ? LAEUFE.filter((nr) => geaendert(nr)) : []);
	const brauchtGrund = $derived(geaenderteLaeufe.some((nr) => laufErfasst(starter?.laeufe[nr])));

	async function speichern(e: SubmitEvent) {
		e.preventDefault();
		if (!starter) return;
		for (const nr of geaenderteLaeufe) {
			const neu = eingabe(nr);
			if (neu === undefined) return ui.melden(`${LAUF_NAMEN[nr]}: Die Zeit ist ungültig.`, 'fehler');
			if (neu && neu.status !== 'ok' && !neu.kommentar) return ui.melden(`${LAUF_NAMEN[nr]}: Für ${(neu.status ?? "").toUpperCase()} ist ein Kommentar erforderlich.`, 'warnung');
		}
		if (brauchtGrund && !grund.trim()) return ui.melden('Bitte den Grund der Änderung angeben.', 'warnung');
		speichernAktiv = true;
		try {
			for (const nr of geaenderteLaeufe) {
				await store.laufSpeichern(starter.id, nr, eingabe(nr)!, grund);
			}
			ui.melden(`Läufe von Nr. ${starter.startnummer} gespeichert.`);
			onschliessen();
		} catch (err) {
			ui.fehler(err, 'Speichern fehlgeschlagen');
		} finally {
			speichernAktiv = false;
		}
	}

	function laufKurz(l: LaufEingabe | null): string {
		if (!l) return 'leer';
		if ((l.status ?? 'ok') !== 'ok') return `${l.status?.toUpperCase()}${l.kommentar ? ` (${l.kommentar})` : ''}`;
		const teile = [`${formatZeit(l.zeit)} s`];
		if (l.fehler1) teile.push(`${l.fehler1}× ${store.v.fehler1Name}`);
		if (l.fehler2) teile.push(`${l.fehler2}× ${store.v.fehler2Name}`);
		return teile.join(', ');
	}
</script>

<Dialog offen={starter !== null} titel={starter ? `Läufe bearbeiten – Nr. ${starter.startnummer} ${anzeigeName(starter)}` : ''} breite="max-w-3xl" {onschliessen}>
	{#if starter && formular[0]}
		<form id="lauf-bearbeiten" onsubmit={speichern} class="flex flex-col gap-3">
			<div class="grid grid-cols-[110px_150px_70px_70px_100px_1fr] items-end gap-2 text-xs font-semibold tracking-wide text-muted uppercase">
				<span>Lauf</span><span>Status</span><span>{store.v.fehler1Name}</span><span>{store.v.fehler2Name}</span><span>Zeit</span><span>Kommentar</span>
			</div>
			{#each LAEUFE as nr (nr)}
				{@const f = formular[nr]}
				{@const neu = eingabe(nr)}
				<div class="grid grid-cols-[110px_150px_70px_70px_100px_1fr] items-center gap-2 rounded-lg {geaendert(nr) ? 'bg-warn-soft/60' : ''} p-1">
					<span class="text-sm font-medium">{LAUF_NAMEN[nr]}</span>
					<select class="input py-1.5" bind:value={f.status} aria-label="Status {LAUF_NAMEN[nr]}">
						<option value="ok">Gefahren</option>
						<option value="dns">DNS</option>
						<option value="dsq">DSQ</option>
					</select>
					<input class="input py-1.5 text-center tabular" type="number" min="0" bind:value={f.fehler1} disabled={f.status !== 'ok'} aria-label="{store.v.fehler1Name} {LAUF_NAMEN[nr]}" />
					<input class="input py-1.5 text-center tabular" type="number" min="0" bind:value={f.fehler2} disabled={f.status !== 'ok'} aria-label="{store.v.fehler2Name} {LAUF_NAMEN[nr]}" />
					<input
						class="input py-1.5 text-right tabular {neu === undefined ? 'border-danger' : ''}"
						inputmode="decimal"
						placeholder="leer"
						bind:value={f.zeitText}
						disabled={f.status !== 'ok'}
						aria-label="Zeit {LAUF_NAMEN[nr]}"
					/>
					<input
						class="input py-1.5 {f.status !== 'ok' && !f.kommentar.trim() ? 'border-warn' : ''}"
						placeholder={f.status !== 'ok' ? 'Pflicht bei DNS/DSQ' : 'optional'}
						bind:value={f.kommentar}
						aria-label="Kommentar {LAUF_NAMEN[nr]}"
					/>
				</div>
				{#if neu && neu.status === 'ok' && neu.zeit !== null}
					<p class="-mt-2 pl-[118px] text-xs text-muted">Ergebnis: {formatZeit(laufErgebnis(neu, store.v))} s</p>
				{/if}
			{/each}

			<div class="mt-2">
				<label class="label" for="aenderungsgrund">
					Grund der Änderung {#if brauchtGrund}<span class="text-danger normal-case">(Pflicht)</span>{/if}
				</label>
				<input id="aenderungsgrund" class="input {brauchtGrund && !grund.trim() ? 'border-warn' : ''}" bind:value={grund} placeholder="z. B. Zeit falsch abgelesen, Protest stattgegeben" />
			</div>
		</form>

		<div class="mt-6">
			<p class="label flex items-center gap-1.5"><History size={13} /> Änderungsprotokoll</p>
			{#if aenderungen.length === 0}
				<p class="text-sm text-muted">Keine nachträglichen Änderungen.</p>
			{:else}
				<ul class="divide-y divide-line rounded-lg border border-line text-sm">
					{#each aenderungen as a (a.id)}
						<li class="px-3 py-2">
							<div class="flex justify-between text-xs text-muted">
								<span class="font-semibold text-fg">{LAUF_NAMEN[a.nr]}</span>
								<span>{new Date(a.zeitpunkt).toLocaleString('de-DE', { dateStyle: 'short', timeStyle: 'medium' })}</span>
							</div>
							<div>{laufKurz(a.vorher)} → <strong>{laufKurz(a.nachher)}</strong></div>
							<div class="text-xs text-muted">Grund: {a.kommentar}</div>
						</li>
					{/each}
				</ul>
			{/if}
		</div>
	{/if}
	{#snippet aktionen()}
		<button class="btn" onclick={onschliessen}>Abbrechen</button>
		<button class="btn btn-primary" type="submit" form="lauf-bearbeiten" disabled={speichernAktiv || geaenderteLaeufe.length === 0}>
			<Save size={16} /> Speichern
		</button>
	{/snippet}
</Dialog>
