<script lang="ts">
	import { onDestroy, onMount, tick, untrack } from 'svelte';
	import { Ban, Eraser, FileClock, ListOrdered, RefreshCw, Save, TriangleAlert } from '@lucide/svelte';
	import {
		anzeigeName,
		LAEUFE,
		LAUF_KURZ,
		LAUF_NAMEN,
		laufErfasst,
		type LaufEingabe,
		type LaufNr,
		type LaufStatus,
		type Starter
	} from '$lib/domain/typen';
	import { naechsterStart, offeneStarts, type StartPlatz } from '$lib/domain/reihenfolge';
	import { laufErgebnis } from '$lib/domain/wertung';
	import { formatZeit, parseZeit } from '$lib/domain/zahlen';
	import type { GemesseneZeit } from '$lib/domain/zeitquelle';
	import { Zeitmessung } from '$lib/stores/zeitmessung.svelte';
	import { istDesktop } from '$lib/plattform';
	import { ui } from '$lib/ui/ui-zustand.svelte';

	let { data } = $props();
	const s = $derived(data.store);
	const store0 = untrack(() => data.store);

	let lauf = $state<LaufNr>(1);
	let nummerText = $state('');
	let status = $state<LaufStatus>('ok');
	let fehler1 = $state<number | null>(0);
	let fehler2 = $state<number | null>(0);
	let zeitText = $state('');
	let importId = $state<string | null>(null);
	let kommentar = $state('');
	let aenderungsgrund = $state('');
	let letzte = $state<{ startId: number; lauf: LaufNr }[]>([]);

	let nummerFeld: HTMLInputElement | undefined = $state();
	let fehler1Feld: HTMLInputElement | undefined = $state();
	let fehler2Feld: HTMLInputElement | undefined = $state();
	let zeitFeld: HTMLInputElement | undefined = $state();
	let kommentarFeld: HTMLInputElement | undefined = $state();
	let grundFeld: HTMLInputElement | undefined = $state();

	const zeitmessung = new Zeitmessung(store0.v.zeitquelle);
	onMount(() => {
		zeitmessung.starten();
		// Mit dem ersten offenen Start der Reihenfolge beginnen
		const erster = naechsterStart(store0.reihenfolge);
		if (erster) platzLaden(erster);
		else nummerFeld?.focus();
	});
	onDestroy(() => zeitmessung.beenden());

	const starter = $derived(nummerText.trim() ? s.nachStartnummer.get(Number(nummerText)) : undefined);
	const vorhanden = $derived(starter?.laeufe[lauf]);
	const vorhandenErfasst = $derived(laufErfasst(vorhanden));
	const zeit = $derived(parseZeit(zeitText));
	const vorschau = $derived(
		status !== 'ok' || zeit === null || zeit === undefined
			? null
			: laufErgebnis({ fehler1: fehler1 ?? 0, fehler2: fehler2 ?? 0, zeit }, s.v)
	);

	/** Formularinhalt als Laufeingabe. */
	const eingabe = $derived<LaufEingabe>({
		status,
		fehler1: status === 'ok' ? (fehler1 ?? 0) : 0,
		fehler2: status === 'ok' ? (fehler2 ?? 0) : 0,
		zeit: status === 'ok' ? (zeit ?? null) : null,
		importId: status === 'ok' ? importId : null,
		kommentar: kommentar.trim() || null
	});

	/** Weicht das Formular von einem bereits erfassten Lauf ab? Dann ist ein Änderungsgrund nötig. */
	const istKorrektur = $derived(
		vorhandenErfasst &&
			!!vorhanden &&
			((vorhanden.status ?? 'ok') !== eingabe.status ||
				(vorhanden.fehler1 || 0) !== eingabe.fehler1 ||
				(vorhanden.fehler2 || 0) !== eingabe.fehler2 ||
				(vorhanden.zeit ?? null) !== (eingabe.zeit ?? null) ||
				(vorhanden.kommentar ?? '') !== (eingabe.kommentar ?? ''))
	);

	const aktuellerPlatz = $derived(starter ? { starterId: starter.id, lauf } : null);
	const alsNaechstes = $derived(offeneStarts(s.reihenfolge, aktuellerPlatz, 10).filter((p) => !(p.starter.id === starter?.id && p.lauf === lauf)));
	const erfasstAnzahl = (nr: LaufNr) => s.starter.filter((st) => laufErfasst(st.laeufe[nr])).length;

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
		status = l?.status ?? 'ok';
		fehler1 = l?.fehler1 ?? 0;
		fehler2 = l?.fehler2 ?? 0;
		zeitText = l?.zeit != null ? formatZeit(l.zeit) : '';
		importId = l?.importId ?? null;
		kommentar = l?.kommentar ?? '';
		aenderungsgrund = '';
	}

	async function platzLaden(p: StartPlatz) {
		lauf = p.lauf;
		nummerText = String(p.starter.startnummer);
		formularLaden(p.starter);
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
		if (status === 'ok') {
			if (zeit === undefined) {
				ui.melden('Die Zeit ist ungültig. Erlaubt sind z. B. 32,45 oder 1:02,34.', 'fehler');
				zeitFeld?.select();
				return;
			}
			if (zeit === null) {
				ui.melden('Bitte eine Zeit eingeben – oder den Lauf als DNS/DSQ kennzeichnen.', 'warnung');
				zeitFeld?.focus();
				return;
			}
		} else if (!kommentar.trim()) {
			ui.melden(`Für ${status.toUpperCase()} ist ein Kommentar erforderlich.`, 'warnung');
			kommentarFeld?.focus();
			return;
		}
		if (istKorrektur && !aenderungsgrund.trim()) {
			ui.melden('Der Lauf war bereits erfasst. Bitte den Grund der Änderung angeben.', 'warnung');
			grundFeld?.focus();
			return;
		}
		const gespeichert = { starterId: starter.id, lauf };
		try {
			await s.laufSpeichern(starter.id, lauf, eingabe, istKorrektur ? aenderungsgrund : undefined);
			letzte = [{ startId: starter.id, lauf }, ...letzte.filter((l) => !(l.startId === starter.id && l.lauf === lauf))].slice(0, 10);
			// Zum nächsten offenen Start in der Startreihenfolge springen
			const naechster = naechsterStart(s.reihenfolge, gespeichert);
			if (naechster) await platzLaden(naechster);
			else {
				ui.melden('Alle Läufe sind erfasst.', 'info');
				await zuruecksetzen();
			}
		} catch (e) {
			ui.fehler(e, 'Speichern fehlgeschlagen');
		}
	}

	async function eingabeLoeschen() {
		if (!starter || !vorhanden) return;
		if (!aenderungsgrund.trim()) {
			ui.melden('Bitte zuerst den Grund für das Löschen angeben.', 'warnung');
			await tick();
			grundFeld?.focus();
			return;
		}
		const ok = await ui.bestaetigen(`${LAUF_NAMEN[lauf]} von Nr. ${starter.startnummer} (${anzeigeName(starter)}) löschen?`, { ja: 'Löschen', gefaehrlich: true });
		if (!ok) return;
		try {
			await s.laufSpeichern(starter.id, lauf, null, aenderungsgrund);
			formularLaden(starter);
		} catch (e) {
			ui.fehler(e);
		}
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
		status = 'ok';
		zeitText = formatZeit(z.zeit);
		importId = z.id;
		await tick();
		(starter ? zeitFeld : nummerFeld)?.focus();
	}

	/** Enter springt zum nächsten Feld, im letzten Feld wird gespeichert. */
	function weiter(e: KeyboardEvent, naechstes: HTMLInputElement | undefined | 'speichern') {
		if (e.key !== 'Enter') return;
		e.preventDefault();
		if (naechstes === 'speichern') speichern();
		else naechstes?.select();
	}

	/** Nach der Zeit: ggf. noch Änderungsgrund, sonst speichern. */
	const nachZeit = $derived(istKorrektur ? grundFeld : ('speichern' as const));
	const nachKommentar = $derived(istKorrektur ? grundFeld : ('speichern' as const));

	async function nummerBestaetigen(e: KeyboardEvent) {
		if (e.key !== 'Enter') return;
		e.preventDefault();
		if (!starter) {
			if (nummerText.trim()) ui.melden(`Startnummer ${nummerText} ist nicht gemeldet.`, 'warnung');
			return;
		}
		// Ist der gewählte Lauf schon erfasst, zum ersten offenen Lauf des Fahrers wechseln
		if (laufErfasst(starter.laeufe[lauf])) {
			const offen = LAEUFE.find((nr) => !laufErfasst(starter.laeufe[nr]));
			if (offen !== undefined) lauf = offen;
		}
		formularLaden(starter);
		await tick();
		(status === 'ok' ? fehler1Feld : kommentarFeld)?.select();
	}

	async function statusSetzen(neu: LaufStatus) {
		status = neu;
		await tick();
		if (neu === 'ok') fehler1Feld?.select();
		else kommentarFeld?.focus();
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

	function laufText(l: LaufEingabe | undefined): string {
		if (!l) return '–';
		if ((l.status ?? 'ok') !== 'ok') return (l.status ?? '').toUpperCase();
		return l.zeit != null ? formatZeit(laufErgebnis(l, s.v)) : '–';
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
					<p class="mt-1 text-center text-xs font-semibold text-accent-strong">{LAUF_NAMEN[lauf]}</p>
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
									<span class="badge {(l?.status ?? 'ok') !== 'ok' ? 'bg-danger-soft text-danger' : laufErfasst(l) ? 'bg-ok-soft text-ok' : 'bg-sunken text-muted'}">
										{LAUF_KURZ[nr]}: {laufText(l)}
									</span>
								{/each}
							</div>
						</div>
					{:else if nummerText.trim()}
						<p class="flex items-center gap-2 text-sm text-danger"><TriangleAlert size={16} /> Startnummer {nummerText} ist nicht gemeldet.</p>
					{:else}
						<p class="text-sm text-muted">Startnummer eingeben und mit ↵ bestätigen – oder rechts einen Start aus der Reihenfolge wählen.</p>
					{/if}
				</div>
			</div>

			<div class="mt-5 flex flex-wrap gap-2" role="group" aria-label="Status">
				<button type="button" class="chip" aria-pressed={status === 'ok'} onclick={() => statusSetzen('ok')} disabled={!starter}>Gefahren</button>
				<button type="button" class="chip" aria-pressed={status === 'dns'} onclick={() => statusSetzen('dns')} disabled={!starter}>
					<Ban size={14} /> DNS – nicht gestartet
				</button>
				<button type="button" class="chip" aria-pressed={status === 'dsq'} onclick={() => statusSetzen('dsq')} disabled={!starter}>
					<Ban size={14} /> DSQ – disqualifiziert
				</button>
			</div>

			{#if status === 'ok'}
				<div class="mt-4 grid grid-cols-2 gap-4 md:grid-cols-[1fr_1fr_1.6fr]">
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
							onkeydown={(e) => weiter(e, nachZeit)}
							disabled={!starter}
						/>
					</div>
				</div>
			{/if}

			<div class="mt-4 grid gap-4 {istKorrektur ? 'md:grid-cols-2' : ''}">
				<div>
					<label class="label" for="kommentar">
						Kommentar {#if status !== 'ok'}<span class="text-danger normal-case">(Pflicht bei {status.toUpperCase()})</span>{:else}<span class="normal-case">(optional)</span>{/if}
					</label>
					<input
						id="kommentar"
						bind:this={kommentarFeld}
						class="input {status !== 'ok' && !kommentar.trim() ? 'border-warn' : ''}"
						placeholder={status === 'dns' ? 'z. B. Fahrer nicht erschienen' : status === 'dsq' ? 'z. B. Frühstart, Streckenabkürzung' : 'Notiz zum Lauf'}
						bind:value={kommentar}
						onkeydown={(e) => weiter(e, nachKommentar)}
						disabled={!starter}
					/>
				</div>
				{#if istKorrektur}
					<div>
						<label class="label" for="grund">Grund der Änderung <span class="text-danger normal-case">(Pflicht)</span></label>
						<input
							id="grund"
							bind:this={grundFeld}
							class="input {aenderungsgrund.trim() ? '' : 'border-warn'}"
							placeholder="z. B. Zeit falsch abgelesen"
							bind:value={aenderungsgrund}
							onkeydown={(e) => weiter(e, 'speichern')}
						/>
					</div>
				{/if}
			</div>

			<div class="mt-6 flex flex-wrap items-center gap-3">
				<div class="mr-auto">
					<p class="text-xs font-semibold tracking-wide text-muted uppercase">Ergebnis {LAUF_NAMEN[lauf]}</p>
					<p class="text-3xl font-bold tabular">{status !== 'ok' ? status.toUpperCase() : vorschau === null ? '–' : `${formatZeit(vorschau)} s`}</p>
					{#if vorhandenErfasst}
						<p class="text-xs text-warn">Bereits erfasst: {laufText(vorhanden)} – Änderungen werden mit Begründung protokolliert.</p>
					{/if}
				</div>
				{#if vorhanden}
					<button type="button" class="btn" onclick={eingabeLoeschen}><Eraser size={16} /> Lauf löschen</button>
				{/if}
				<button type="button" class="btn" onclick={zuruecksetzen}>Abbrechen <kbd class="text-xs opacity-60">Esc</kbd></button>
				<button type="submit" class="btn btn-primary px-5 py-3" disabled={!starter}><Save size={16} /> Speichern & weiter <kbd class="text-xs opacity-70">↵</kbd></button>
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
								<button class="flex w-full items-center gap-4 px-5 py-2 text-left hover:bg-sunken" onclick={() => platzLaden({ starter: st, lauf: l.lauf })}>
									<span class="w-12 font-bold tabular">{st.startnummer}</span>
									<span class="flex-1">{anzeigeName(st)}</span>
									<span class="text-muted">{LAUF_KURZ[l.lauf]}</span>
									<span class="w-24 text-right font-semibold tabular">{laufText(st.laeufe[l.lauf])}</span>
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
			<h2 class="flex items-center gap-2 border-b border-line px-4 py-3 text-sm font-semibold"><ListOrdered size={16} /> Als Nächstes</h2>
			{#if alsNaechstes.length === 0}
				<p class="px-4 py-3 text-sm text-ok">Keine weiteren offenen Starts.</p>
			{:else}
				<ul class="divide-y divide-line text-sm">
					{#each alsNaechstes as p (p.starter.id + '-' + p.lauf)}
						<li>
							<button class="flex w-full items-center gap-3 px-4 py-1.5 text-left hover:bg-sunken" onclick={() => platzLaden(p)}>
								<span class="w-10 font-bold tabular">{p.starter.startnummer}</span>
								<span class="min-w-0 flex-1 truncate">{anzeigeName(p.starter)}</span>
								<span class="badge {p.lauf === 0 ? 'bg-sunken text-muted' : 'bg-accent-soft text-accent-strong'}">{LAUF_KURZ[p.lauf]}</span>
							</button>
						</li>
					{/each}
				</ul>
			{/if}
			<p class="border-t border-line px-4 py-2 text-[11px] text-muted">Reihenfolge: je zwei Fahrer Training und Wertung 1, danach alle Wertung 2.</p>
		</section>

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
	</aside>
</div>
