<script lang="ts">
	import { onDestroy, onMount, tick, untrack } from 'svelte';
	import { Eraser, FileClock, RefreshCw, Save, TriangleAlert } from '@lucide/svelte';
	import { anzeigeName, LAEUFE, LAUF_KURZ, LAUF_NAMEN, type LaufNr, type Starter } from '$lib/domain/typen';
	import { laufErgebnis } from '$lib/domain/wertung';
	import { formatZeit, parseZeit } from '$lib/domain/zahlen';
	import type { GemesseneZeit } from '$lib/domain/zeitquelle';
	import { Zeitmessung } from '$lib/stores/zeitmessung.svelte';
	import { istDesktop } from '$lib/plattform';
	import { ui } from '$lib/ui/ui-zustand.svelte';

	let { data } = $props();
	const s = $derived(data.store);

	const store0 = untrack(() => data.store);
	const LAUF_SCHLUESSEL = `erfassung-lauf-${store0.id}`;
	function gemerkterLauf(): LaufNr {
		try {
			const gespeichert = sessionStorage.getItem(LAUF_SCHLUESSEL);
			const wert = gespeichert === null ? NaN : Number(gespeichert);
			return (LAEUFE as number[]).includes(wert) ? (wert as LaufNr) : 1;
		} catch {
			return 1;
		}
	}

	let lauf = $state<LaufNr>(gemerkterLauf());
	let nummerText = $state('');
	let fehler1 = $state<number | null>(0);
	let fehler2 = $state<number | null>(0);
	let zeitText = $state('');
	let importId = $state<string | null>(null);
	let letzte = $state<{ startId: number; lauf: LaufNr; ergebnis: number | null }[]>([]);

	let nummerFeld: HTMLInputElement | undefined = $state();
	let fehler1Feld: HTMLInputElement | undefined = $state();
	let fehler2Feld: HTMLInputElement | undefined = $state();
	let zeitFeld: HTMLInputElement | undefined = $state();

	const zeitmessung = new Zeitmessung(store0.v.zeitquelle);
	onMount(() => {
		zeitmessung.starten();
		nummerFeld?.focus();
	});
	onDestroy(() => zeitmessung.beenden());

	$effect(() => {
		try {
			sessionStorage.setItem(LAUF_SCHLUESSEL, String(lauf));
		} catch {
			/* ignorieren */
		}
	});

	const starter = $derived(nummerText.trim() ? s.nachStartnummer.get(Number(nummerText)) : undefined);
	const vorhanden = $derived(starter?.laeufe[lauf]);
	const zeit = $derived(parseZeit(zeitText));
	const vorschau = $derived(
		zeit === null || zeit === undefined ? null : laufErgebnis({ fehler1: fehler1 ?? 0, fehler2: fehler2 ?? 0, zeit }, s.v)
	);
	const offen = $derived(s.starter.filter((st) => st.laeufe[lauf]?.zeit == null));
	const erfasstAnzahl = (nr: LaufNr) => s.starter.filter((st) => st.laeufe[nr]?.zeit != null).length;

	/** Welche Kennungen der Zeitmessung sind schon einem Lauf zugeordnet? */
	const zugeordnet = $derived.by(() => {
		const map = new Map<string, { starter: Starter; lauf: LaufNr }>();
		for (const st of s.starter)
			for (const nr of LAEUFE) {
				const id = st.laeufe[nr]?.importId;
				if (id) map.set(id, { starter: st, lauf: nr });
			}
		return map;
	});
	const naechsteFreieZeit = $derived(zeitmessung.zeiten.find((z) => !zugeordnet.has(z.id)));

	function formularLaden(st: Starter | undefined) {
		const l = st?.laeufe[lauf];
		fehler1 = l?.fehler1 ?? 0;
		fehler2 = l?.fehler2 ?? 0;
		zeitText = l?.zeit != null ? formatZeit(l.zeit) : '';
		importId = l?.importId ?? null;
	}

	async function starterWaehlen(st: Starter) {
		nummerText = String(st.startnummer);
		formularLaden(st);
		await tick();
		fehler1Feld?.select();
	}

	async function zuruecksetzen() {
		nummerText = '';
		formularLaden(undefined);
		await tick();
		nummerFeld?.focus();
	}

	async function laufWechseln(nr: LaufNr) {
		lauf = nr;
		formularLaden(starter);
		await tick();
		(starter ? fehler1Feld : nummerFeld)?.select();
	}

	async function speichern() {
		if (!starter) {
			ui.melden('Bitte zuerst eine gültige Startnummer eingeben.', 'warnung');
			nummerFeld?.select();
			return;
		}
		if (zeit === undefined) {
			ui.melden('Die Zeit ist ungültig. Erlaubt sind z. B. 32,45 oder 1:02,34.', 'fehler');
			zeitFeld?.select();
			return;
		}
		if (zeit === null) {
			ui.melden('Bitte eine Zeit eingeben.', 'warnung');
			zeitFeld?.focus();
			return;
		}
		try {
			await s.laufSpeichern(starter.id, lauf, { fehler1: fehler1 ?? 0, fehler2: fehler2 ?? 0, zeit, importId });
			letzte = [{ startId: starter.id, lauf, ergebnis: vorschau }, ...letzte.filter((l) => !(l.startId === starter.id && l.lauf === lauf))].slice(0, 10);
			await zuruecksetzen();
		} catch (e) {
			ui.fehler(e, 'Speichern fehlgeschlagen');
		}
	}

	async function eingabeLoeschen() {
		if (!starter || !vorhanden) return;
		const ok = await ui.bestaetigen(`${LAUF_NAMEN[lauf]} von Nr. ${starter.startnummer} (${anzeigeName(starter)}) löschen?`, { ja: 'Löschen', gefaehrlich: true });
		if (!ok) return;
		await s.laufSpeichern(starter.id, lauf, null);
		await zuruecksetzen();
	}

	async function zeitUebernehmen(z: GemesseneZeit | undefined) {
		if (!z) {
			ui.melden('Keine freie Zeit in der Zeitmessung vorhanden.', 'info');
			return;
		}
		const belegt = zugeordnet.get(z.id);
		if (belegt && belegt.starter.id !== starter?.id) {
			const ok = await ui.bestaetigen(
				`Die Zeit mit Kennung ${z.id} ist bereits Nr. ${belegt.starter.startnummer} (${LAUF_KURZ[belegt.lauf]}) zugeordnet. Trotzdem übernehmen?`,
				{ ja: 'Übernehmen' }
			);
			if (!ok) return;
		}
		if (vorhanden?.zeit != null && vorhanden.zeit !== z.zeit) {
			const ok = await ui.bestaetigen(`Für diesen Lauf ist bereits ${formatZeit(vorhanden.zeit)} s erfasst. Durch ${formatZeit(z.zeit)} s ersetzen?`, { ja: 'Ersetzen' });
			if (!ok) return;
		}
		zeitText = formatZeit(z.zeit);
		importId = z.id;
		await tick();
		(starter ? zeitFeld : nummerFeld)?.focus();
	}

	/** Enter springt zum nächsten Feld, im Zeitfeld wird gespeichert. */
	function weiter(e: KeyboardEvent, naechstes: HTMLInputElement | undefined | 'speichern') {
		if (e.key !== 'Enter') return;
		e.preventDefault();
		if (naechstes === 'speichern') speichern();
		else naechstes?.select();
	}

	async function nummerBestaetigen(e: KeyboardEvent) {
		if (e.key !== 'Enter') return;
		e.preventDefault();
		if (!starter) {
			if (nummerText.trim()) ui.melden(`Startnummer ${nummerText} ist nicht gemeldet.`, 'warnung');
			return;
		}
		await starterWaehlen(starter);
	}

	function globaleTasten(e: KeyboardEvent) {
		if (e.key === 'Escape' && !document.querySelector('dialog[open]')) {
			zuruecksetzen();
			return;
		}
		if (!(e.ctrlKey || e.metaKey)) return;
		if (['0', '1', '2'].includes(e.key)) {
			e.preventDefault();
			laufWechseln(Number(e.key) as LaufNr);
		} else if (e.key.toLowerCase() === 't') {
			e.preventDefault();
			zeitUebernehmen(naechsteFreieZeit);
		} else if (e.key.toLowerCase() === 's') {
			e.preventDefault();
			speichern();
		}
	}
</script>

<svelte:window onkeydown={globaleTasten} />

<div class="grid gap-6 px-8 py-6 xl:grid-cols-[minmax(0,1fr)_360px]">
	<div class="flex min-w-0 flex-col gap-6">
		<div class="flex flex-wrap gap-2" role="group" aria-label="Lauf">
			{#each LAEUFE as nr (nr)}
				<button class="chip" aria-pressed={lauf === nr} onclick={() => laufWechseln(nr)}>
					{LAUF_NAMEN[nr]}
					<span class="text-xs opacity-70 tabular">{erfasstAnzahl(nr)}/{s.starter.length}</span>
					<kbd class="rounded border border-current/30 px-1 text-[10px] opacity-60">Strg+{nr}</kbd>
				</button>
			{/each}
		</div>

		<form class="card p-6" onsubmit={(e) => (e.preventDefault(), speichern())}>
			<div class="grid gap-6 md:grid-cols-[180px_minmax(0,1fr)]">
				<div>
					<label class="label" for="nummer">Startnummer</label>
					<input
						id="nummer"
						bind:this={nummerFeld}
						class="input py-3 text-center text-3xl font-bold tabular"
						inputmode="numeric"
						autocomplete="off"
						bind:value={nummerText}
						oninput={() => formularLaden(starter)}
						onkeydown={nummerBestaetigen}
					/>
				</div>
				<div class="flex min-h-24 items-center rounded-xl border border-dashed border-line px-5 py-3">
					{#if starter}
						{@const klasse = s.klasseVon(starter)}
						<div class="min-w-0">
							<p class="truncate text-xl font-bold">{anzeigeName(starter)}</p>
							<p class="mt-0.5 text-sm text-muted">{[klasse?.name, starter.verein].filter(Boolean).join(' · ')}</p>
							<div class="mt-1.5 flex flex-wrap gap-1.5">
								{#if starter.rookieJahr !== null && starter.rookieJahr === s.jahr}<span class="badge bg-info-soft text-info">Rookie</span>{/if}
								{#if starter.ausserWertung}<span class="badge bg-warn-soft text-warn">außer Wertung</span>{/if}
								{#each LAEUFE as nr (nr)}
									{@const l = starter.laeufe[nr]}
									<span class="badge {l?.zeit != null ? 'bg-ok-soft text-ok' : 'bg-sunken text-muted'}">
										{LAUF_KURZ[nr]}: {l?.zeit != null ? formatZeit(laufErgebnis(l, s.v)) : '–'}
									</span>
								{/each}
							</div>
						</div>
					{:else if nummerText.trim()}
						<p class="flex items-center gap-2 text-sm text-danger"><TriangleAlert size={16} /> Startnummer {nummerText} ist nicht gemeldet.</p>
					{:else}
						<p class="text-sm text-muted">Startnummer eingeben und mit ↵ bestätigen – oder rechts einen offenen Fahrer anklicken.</p>
					{/if}
				</div>
			</div>

			<div class="mt-6 grid grid-cols-2 gap-4 md:grid-cols-[1fr_1fr_1.6fr]">
				<div>
					<label class="label" for="f1">{s.v.fehler1Name} <span class="normal-case">(× {s.v.strafe1} s)</span></label>
					<input id="f1" bind:this={fehler1Feld} class="input py-3 text-center text-2xl font-semibold tabular" type="number" min="0" bind:value={fehler1} onkeydown={(e) => weiter(e, fehler2Feld)} disabled={!starter} />
				</div>
				<div>
					<label class="label" for="f2">{s.v.fehler2Name} <span class="normal-case">(× {s.v.strafe2} s)</span></label>
					<input id="f2" bind:this={fehler2Feld} class="input py-3 text-center text-2xl font-semibold tabular" type="number" min="0" bind:value={fehler2} onkeydown={(e) => weiter(e, zeitFeld)} disabled={!starter} />
				</div>
				<div class="col-span-2 md:col-span-1">
					<label class="label" for="zeit">Zeit in Sekunden {#if importId}<span class="normal-case">· Zeitmessung #{importId}</span>{/if}</label>
					<input
						id="zeit"
						bind:this={zeitFeld}
						class="input py-3 text-center text-2xl font-semibold tabular {zeit === undefined ? 'border-danger' : ''}"
						inputmode="decimal"
						autocomplete="off"
						placeholder="0,00"
						bind:value={zeitText}
						oninput={() => (importId = null)}
						onkeydown={(e) => weiter(e, 'speichern')}
						disabled={!starter}
					/>
				</div>
			</div>

			<div class="mt-6 flex flex-wrap items-center gap-3">
				<div class="mr-auto">
					<p class="text-xs font-semibold tracking-wide text-muted uppercase">Ergebnis {LAUF_NAMEN[lauf]}</p>
					<p class="text-3xl font-bold tabular">{vorschau === null ? '–' : `${formatZeit(vorschau)} s`}</p>
					{#if vorhanden?.zeit != null}
						<p class="text-xs text-warn">Bereits erfasst: {formatZeit(laufErgebnis(vorhanden, s.v))} s – Speichern überschreibt.</p>
					{/if}
				</div>
				{#if vorhanden}
					<button type="button" class="btn" onclick={eingabeLoeschen}><Eraser size={16} /> Lauf löschen</button>
				{/if}
				<button type="button" class="btn" onclick={zuruecksetzen}>Abbrechen <kbd class="text-xs opacity-60">Esc</kbd></button>
				<button type="submit" class="btn btn-primary px-5 py-3" disabled={!starter}><Save size={16} /> Speichern <kbd class="text-xs opacity-70">↵</kbd></button>
			</div>
		</form>

		{#if letzte.length}
			<section class="card overflow-hidden">
				<h2 class="border-b border-line px-5 py-3 text-sm font-semibold">Zuletzt erfasst</h2>
				<ul class="divide-y divide-line text-sm">
					{#each letzte as l (l.startId + '-' + l.lauf)}
						{@const st = s.starter.find((x) => x.id === l.startId)}
						{#if st}
							<li>
								<button class="flex w-full items-center gap-4 px-5 py-2 text-left hover:bg-sunken" onclick={() => (laufWechseln(l.lauf), starterWaehlen(st))}>
									<span class="w-12 font-bold tabular">{st.startnummer}</span>
									<span class="flex-1">{anzeigeName(st)}</span>
									<span class="text-muted">{LAUF_KURZ[l.lauf]}</span>
									<span class="w-24 text-right font-semibold tabular">{formatZeit(laufErgebnis(st.laeufe[l.lauf], s.v))}</span>
								</button>
							</li>
						{/if}
					{/each}
				</ul>
			</section>
		{/if}
	</div>

	<aside class="flex flex-col gap-6">
		<section class="card overflow-hidden">
			<header class="flex items-center justify-between border-b border-line px-4 py-3">
				<h2 class="flex items-center gap-2 text-sm font-semibold"><FileClock size={16} /> Zeitmessung</h2>
				{#if zeitmessung.konfiguriert || !istDesktop()}
					<button class="btn btn-sm" onclick={() => (istDesktop() ? zeitmessung.einlesen() : zeitmessung.dateiWaehlen())}>
						<RefreshCw size={14} />
						{istDesktop() ? 'Neu einlesen' : 'Datei laden'}
					</button>
				{/if}
			</header>
			{#if !zeitmessung.konfiguriert && istDesktop()}
				<p class="px-4 py-3 text-sm text-muted">
					Keine Zeitmessung eingerichtet. In den <a class="text-accent-strong underline" href="/veranstaltung/{s.id}/einstellungen">Einstellungen</a> kann eine CSV- oder Excel-Datei der Zeitmessanlage hinterlegt werden.
				</p>
			{:else}
				<div class="px-4 py-2 text-xs text-muted">
					{#if zeitmessung.fehler}
						<span class="text-danger">{zeitmessung.fehler}</span>
					{:else if zeitmessung.stand}
						{zeitmessung.dateiname} · {zeitmessung.zeiten.length} Zeiten · Stand {zeitmessung.stand.toLocaleTimeString('de-DE')}
						{#if istDesktop()}<span class="badge ml-1 bg-ok-soft text-ok">live</span>{/if}
					{:else}
						Noch nicht eingelesen.
					{/if}
				</div>
				{#if zeitmessung.zeiten.length}
					<button class="btn btn-primary btn-sm mx-4 mb-2 w-[calc(100%-2rem)]" onclick={() => zeitUebernehmen(naechsteFreieZeit)} disabled={!naechsteFreieZeit}>
						Nächste freie Zeit übernehmen <kbd class="text-[10px] opacity-70">Strg+T</kbd>
					</button>
					<ul class="max-h-80 divide-y divide-line overflow-y-auto border-t border-line text-sm">
						{#each [...zeitmessung.zeiten].reverse().slice(0, 50) as z (z.zeile)}
							{@const belegt = zugeordnet.get(z.id)}
							<li>
								<button class="flex w-full items-center gap-3 px-4 py-1.5 text-left hover:bg-sunken {belegt ? 'opacity-55' : ''}" onclick={() => zeitUebernehmen(z)}>
									<span class="w-14 font-mono text-xs text-muted">#{z.id}</span>
									<span class="flex-1 font-semibold tabular">{formatZeit(z.zeit)}</span>
									{#if belegt}<span class="text-xs text-muted">Nr. {belegt.starter.startnummer} · {LAUF_KURZ[belegt.lauf]}</span>{:else}<span class="badge bg-accent-soft text-accent-strong">frei</span>{/if}
								</button>
							</li>
						{/each}
					</ul>
				{/if}
			{/if}
		</section>

		<section class="card overflow-hidden">
			<h2 class="border-b border-line px-4 py-3 text-sm font-semibold">Offen in {LAUF_NAMEN[lauf]} <span class="font-normal text-muted">({offen.length})</span></h2>
			{#if offen.length === 0}
				<p class="px-4 py-3 text-sm text-ok">Alle Fahrer sind erfasst.</p>
			{:else}
				<div class="flex max-h-96 flex-wrap gap-1.5 overflow-y-auto p-3">
					{#each offen as st (st.id)}
						<button class="rounded-lg border border-line px-2.5 py-1 text-sm font-semibold tabular hover:border-accent hover:bg-accent-soft" title="{anzeigeName(st)} · {s.klasseVon(st)?.name}" onclick={() => starterWaehlen(st)}>
							{st.startnummer}
						</button>
					{/each}
				</div>
			{/if}
		</section>
	</aside>
</div>
