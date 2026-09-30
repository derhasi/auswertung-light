<script lang="ts">
	import { onDestroy, onMount, tick, untrack } from 'svelte';
	import { goto } from '$app/navigation';
	import { ArrowRight, Ban, Eraser, FileClock, Flag, ListOrdered, RefreshCw, Trophy, TriangleAlert } from '@lucide/svelte';
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
	import { klasseKomplett, naechsterStart, offeneStarts, type StartPlatz } from '$lib/domain/reihenfolge';
	import { laufErgebnis } from '$lib/domain/wertung';
	import { formatZeit, parseZeit } from '$lib/domain/zahlen';
	import type { Klasse } from '$lib/domain/typen';
	import Dialog from '$lib/ui/Dialog.svelte';
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
	/** Zwischenschritt nach dem letzten Lauf einer Klasse. */
	let klassenAbschluss = $state<{ klasse: Klasse; naechste: StartPlatz | null } | null>(null);

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
		const klasseId = starter.klasseId;
		const klasseVorherKomplett = klasseKomplett(s.starter, klasseId);
		try {
			await s.laufSpeichern(starter.id, lauf, eingabe, istKorrektur ? aenderungsgrund : undefined);
			letzte = [{ startId: starter.id, lauf }, ...letzte.filter((l) => !(l.startId === starter.id && l.lauf === lauf))].slice(0, 10);
			// Letzter Lauf der Klasse: Zwischenschritt statt direkt weiter
			if (!klasseVorherKomplett && klasseKomplett(s.starter, klasseId)) {
				const klasse = s.klassen.find((k) => k.id === klasseId)!;
				klassenAbschluss = { klasse, naechste: naechsterStart(s.reihenfolge, gespeichert) };
				nummerText = '';
				formularLaden(undefined);
				return;
			}
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
		if (klassenAbschluss || document.querySelector('dialog[open]')) return;
		if (e.key === 'Escape') {
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

	async function naechsteKlasse() {
		const naechste = klassenAbschluss?.naechste;
		klassenAbschluss = null;
		if (naechste) await platzLaden(naechste);
		else await zuruecksetzen();
	}

	const abschlussZeilen = $derived(
		klassenAbschluss ? (s.klassenWertungen.find((k) => k.klasse.id === klassenAbschluss!.klasse.id)?.zeilen ?? []) : []
	);

	function laufText(l: LaufEingabe | undefined): string {
		if (!l) return '–';
		if ((l.status ?? 'ok') !== 'ok') return (l.status ?? '').toUpperCase();
		return l.zeit != null ? formatZeit(laufErgebnis(l, s.v)) : '–';
	}

	/** Fortschrittsleiste je Klasse (alle drei Läufe). */
	const strecke = $derived.by(() => {
		const aktiv = starter?.klasseId;
		return s.klassenWertungen
			.filter((k) => k.zeilen.length)
			.map(({ klasse, zeilen }) => {
				const gesamt = zeilen.length * 3;
				const erfasst = zeilen.reduce((a, z) => a + LAEUFE.filter((nr) => laufErfasst(z.starter.laeufe[nr])).length, 0);
				return { klasse, gesamt, erfasst, fertig: erfasst === gesamt, aktiv: klasse.id === aktiv };
			});
	});
</script>

<svelte:window onkeydown={globaleTasten} />

<div class="flex flex-wrap items-center gap-x-6 gap-y-3 border-b border-line bg-surface px-7 py-4">
	<div class="flex gap-1.5" role="group" aria-label="Lauf">
		{#each LAEUFE as nr (nr)}
			<button class="chip" aria-pressed={lauf === nr} onclick={() => laufWechseln(nr)} title="Strg+{nr}">
				{LAUF_NAMEN[nr]}
				<span class="text-[13px] font-semibold opacity-80 tabular">{erfasstAnzahl(nr)}/{s.starter.length}</span>
			</button>
		{/each}
	</div>
	<div class="flex min-w-80 flex-1 gap-1.5" aria-label="Fortschritt je Klasse">
		{#each strecke as k (k.klasse.id)}
			<div class="flex min-w-0 flex-1 flex-col gap-1.5">
				<div class="h-2 rounded bg-sunken">
					<div class="h-2 rounded {k.fertig ? 'bg-fg' : k.aktiv ? 'bg-accent' : 'bg-muted/40'}" style:width="{(k.erfasst / k.gesamt) * 100}%"></div>
				</div>
				<div class="flex justify-between gap-1 text-[13px] font-semibold {k.aktiv ? 'text-accent' : 'text-muted'}">
					<span class="truncate">{k.klasse.name}</span><span class="tabular">{k.fertig ? 'fertig' : `${k.erfasst}/${k.gesamt}`}</span>
				</div>
			</div>
		{/each}
	</div>
</div>

<div class="grid gap-6 px-7 py-6 xl:grid-cols-[minmax(0,1fr)_340px]">
	<form class="flex min-w-0 flex-col gap-4" onsubmit={(e) => (e.preventDefault(), speichern())}>
		<div class="card flex overflow-hidden">
			<div class="flex w-44 shrink-0 flex-col items-center justify-center gap-0.5 bg-ink py-3 text-on-ink">
				<label class="eyebrow text-xs text-ink-muted" for="nummer">Startnr.</label>
				<input
					id="nummer"
					bind:this={nummerFeld}
					class="display w-36 bg-transparent text-center text-[96px] leading-none text-on-ink tabular outline-none placeholder:text-ink-line"
					inputmode="numeric"
					autocomplete="off"
					placeholder="–"
					bind:value={nummerText}
					oninput={() => formularLaden(starter)}
					onkeydown={nummerBestaetigen}
				/>
			</div>
			<div class="flex min-w-0 flex-1 flex-col justify-center gap-2 px-7 py-5">
				{#if starter}
					{@const klasse = s.klasseVon(starter)}
					<div class="flex flex-wrap gap-2">
						<span class="badge bg-accent px-2.5 py-1 text-[13px] text-on-accent">{LAUF_NAMEN[lauf]}</span>
						{#if klasse}<span class="badge bg-sunken px-2.5 py-1 text-[13px] text-fg">{klasse.name}</span>{/if}
						{#if starter.rookieJahr !== null && starter.rookieJahr === s.jahr}<span class="badge bg-accent-soft px-2.5 py-1 text-[13px] text-accent-strong">Rookie</span>{/if}
						{#if starter.ausserWertung}<span class="badge bg-warn-soft px-2.5 py-1 text-[13px] text-warn">außer Wertung</span>{/if}
					</div>
					<p class="display truncate text-[52px] leading-none">{starter.vorname} {starter.nachname}</p>
					<p class="flex flex-wrap gap-x-3 text-base text-muted">
						{#if starter.verein}<span>{starter.verein}</span>{/if}
						{#each LAEUFE as nr (nr)}
							{@const l = starter.laeufe[nr]}
							<span class="tabular {(l?.status ?? 'ok') !== 'ok' ? 'text-danger' : laufErfasst(l) ? 'text-fg' : ''}">{LAUF_KURZ[nr]}: {laufText(l)}</span>
						{/each}
					</p>
				{:else if nummerText.trim()}
					<p class="flex items-center gap-2 text-lg font-semibold text-danger"><TriangleAlert size={20} /> Startnummer {nummerText} ist nicht gemeldet.</p>
				{:else}
					<p class="display text-[34px] leading-none text-muted">Startnummer eingeben</p>
					<p class="text-base text-muted">Mit ↵ bestätigen – oder rechts einen Start aus der Reihenfolge wählen.</p>
				{/if}
			</div>
		</div>

		<div class="flex flex-wrap gap-2" role="group" aria-label="Status">
			<button type="button" class="chip" aria-pressed={status === 'ok'} onclick={() => statusSetzen('ok')} disabled={!starter}>Gefahren</button>
			<button type="button" class="chip" aria-pressed={status === 'dns'} onclick={() => statusSetzen('dns')} disabled={!starter}><Ban size={15} /> DNS – nicht gestartet</button>
			<button type="button" class="chip" aria-pressed={status === 'dsq'} onclick={() => statusSetzen('dsq')} disabled={!starter}><Ban size={15} /> DSQ – disqualifiziert</button>
		</div>

		{#if status === 'ok'}
			<div class="grid grid-cols-2 gap-3.5 md:grid-cols-3">
				<div class="card flex flex-col gap-2 p-4 focus-within:border-accent focus-within:ring-2 focus-within:ring-accent">
					<label class="flex justify-between text-sm font-bold tracking-wide text-muted uppercase" for="f1">{s.v.fehler1Name} <span class="normal-case opacity-80">× {s.v.strafe1} s</span></label>
					<input id="f1" bind:this={fehler1Feld} class="display w-full bg-transparent text-[72px] leading-none tabular outline-none disabled:opacity-40" type="number" min="0" bind:value={fehler1} onkeydown={(e) => weiter(e, fehler2Feld)} disabled={!starter} />
				</div>
				<div class="card flex flex-col gap-2 p-4 focus-within:border-accent focus-within:ring-2 focus-within:ring-accent">
					<label class="flex justify-between text-sm font-bold tracking-wide text-muted uppercase" for="f2">{s.v.fehler2Name} <span class="normal-case opacity-80">× {s.v.strafe2} s</span></label>
					<input id="f2" bind:this={fehler2Feld} class="display w-full bg-transparent text-[72px] leading-none tabular outline-none disabled:opacity-40" type="number" min="0" bind:value={fehler2} onkeydown={(e) => weiter(e, zeitFeld)} disabled={!starter} />
				</div>
				<div class="card col-span-2 flex flex-col gap-2 p-4 focus-within:border-accent focus-within:ring-2 focus-within:ring-accent md:col-span-1 {zeit === undefined ? 'border-danger' : ''}">
					<label class="flex justify-between text-sm font-bold tracking-wide text-muted uppercase" for="zeit">Zeit <span class="opacity-80">{importId ? `Zeitmessung #${importId}` : 'Sekunden'}</span></label>
					<input
						id="zeit"
						bind:this={zeitFeld}
						class="display w-full bg-transparent text-[72px] leading-none tabular outline-none placeholder:text-muted/40 disabled:opacity-40"
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

		<div class="grid gap-3.5 {istKorrektur ? 'md:grid-cols-2' : ''}">
			<div>
				<label class="label" for="kommentar">
					Kommentar {#if status !== 'ok'}<span class="text-danger normal-case">(Pflicht bei {status.toUpperCase()})</span>{:else}<span class="font-semibold normal-case">(optional)</span>{/if}
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

		<div class="card-ink flex flex-wrap items-center gap-5 px-6 py-4">
			<div class="flex flex-col">
				<span class="eyebrow text-ink-muted">Ergebnis {LAUF_NAMEN[lauf]}</span>
				<span class="display text-[60px] leading-none normal-case tabular">{status !== 'ok' ? status.toUpperCase() : vorschau === null ? '–' : `${formatZeit(vorschau)} s`}</span>
			</div>
			{#if vorhandenErfasst}
				<p class="max-w-64 text-sm text-orange-300">Bereits erfasst: {laufText(vorhanden)}. Änderungen werden mit Begründung protokolliert.</p>
			{/if}
			<div class="ml-auto flex flex-wrap gap-2.5">
				{#if vorhanden}
					<button type="button" class="btn border-ink-line bg-transparent text-on-ink hover:bg-white/10" onclick={eingabeLoeschen}><Eraser size={16} /> Lauf löschen</button>
				{/if}
				<button type="button" class="btn border-ink-line bg-transparent text-on-ink hover:bg-white/10" onclick={zuruecksetzen}>Abbrechen <kbd class="text-xs opacity-60">Esc</kbd></button>
				<button type="submit" class="btn btn-primary btn-lg" disabled={!starter}>Speichern ↵</button>
			</div>
		</div>

		{#if letzte.length}
			<section class="card overflow-hidden">
				<h2 class="section-title border-b border-line px-5 py-3">Zuletzt erfasst</h2>
				<ul class="divide-y divide-line">
					{#each letzte as l (l.startId + '-' + l.lauf)}
						{@const st = s.starter.find((x) => x.id === l.startId)}
						{#if st}
							<li>
								<button type="button" class="flex w-full items-center gap-4 px-5 py-2.5 text-left hover:bg-sunken" onclick={() => platzLaden({ starter: st, lauf: l.lauf })}>
									<span class="display w-12 text-2xl tabular">{st.startnummer}</span>
									<span class="flex-1 font-semibold">{anzeigeName(st)}</span>
									<span class="text-sm text-muted">{LAUF_NAMEN[l.lauf]}</span>
									<span class="display w-24 text-right text-2xl tabular">{laufText(st.laeufe[l.lauf])}</span>
								</button>
							</li>
						{/if}
					{/each}
				</ul>
			</section>
		{/if}
	</form>

	<aside class="flex flex-col gap-3.5">
		<div class="flex items-center justify-between">
			<h2 class="section-title flex items-center gap-2"><ListOrdered size={20} /> Als Nächstes</h2>
			{#if starter}<span class="text-sm text-muted">{s.klasseVon(starter)?.name}</span>{/if}
		</div>
		{#if alsNaechstes.length === 0}
			<p class="card px-4 py-3 text-ok">Keine weiteren offenen Starts.</p>
		{:else}
			<ul class="flex flex-col gap-2">
				{#each alsNaechstes.slice(0, 6) as p (p.starter.id + '-' + p.lauf)}
					<li>
						<button class="card flex w-full items-center gap-3.5 px-3.5 py-2.5 text-left hover:border-accent" onclick={() => platzLaden(p)}>
							<span class="display flex size-11 shrink-0 items-center justify-center rounded-lg bg-sunken text-2xl tabular">{p.starter.startnummer}</span>
							<span class="flex min-w-0 flex-1 flex-col">
								<span class="truncate font-bold">{p.starter.vorname} {p.starter.nachname}</span>
								<span class="truncate text-[13px] text-muted">{p.starter.verein || s.klasseVon(p.starter)?.name}</span>
							</span>
							<span class="text-[13px] font-semibold {p.lauf === 0 ? 'text-muted' : 'text-accent'}">{LAUF_NAMEN[p.lauf]}</span>
						</button>
					</li>
				{/each}
			</ul>
		{/if}
		<p class="text-xs text-muted">Reihenfolge je Klasse: zwei Fahrer Training und Wertungslauf 1, danach alle Wertungslauf 2.</p>

		<section class="mt-2 overflow-hidden rounded-[14px] border border-orange-300 bg-accent-soft">
			<header class="flex items-center justify-between px-4 pt-3">
				<h2 class="eyebrow flex items-center gap-1.5 text-accent-strong"><FileClock size={15} /> Zeitmessung</h2>
				{#if zeitmessung.konfiguriert || !istDesktop()}
					<button class="btn btn-sm border-transparent bg-transparent hover:bg-surface/60" onclick={() => (istDesktop() ? zeitmessung.einlesen() : zeitmessung.dateiWaehlen())}>
						<RefreshCw size={14} />
						{istDesktop() ? 'Neu einlesen' : 'Datei laden'}
					</button>
				{/if}
			</header>
			{#if !zeitmessung.konfiguriert && istDesktop()}
				<p class="px-4 pt-1 pb-4 text-sm">
					Keine Zeitmessung eingerichtet. In den <a class="font-semibold text-accent-strong underline" href="/veranstaltung/{s.id}/einstellungen">Einstellungen</a> kann eine CSV- oder Excel-Datei der Zeitmessanlage hinterlegt werden.
				</p>
			{:else}
				<div class="flex items-center gap-3 px-4 pt-1 pb-3">
					<div class="flex min-w-0 flex-1 flex-col">
						<span class="display text-[30px] leading-tight tabular">{naechsteFreieZeit ? formatZeit(naechsteFreieZeit.zeit) : '–'}</span>
						<span class="truncate text-xs text-muted">
							{#if zeitmessung.fehler}<span class="text-danger">{zeitmessung.fehler}</span>
							{:else if zeitmessung.stand}{zeitmessung.dateiname} · Stand {zeitmessung.stand.toLocaleTimeString('de-DE')}{istDesktop() ? ' · live' : ''}
							{:else}Noch nicht eingelesen.{/if}
						</span>
					</div>
					<button class="btn btn-ink shrink-0" onclick={() => zeitUebernehmen(naechsteFreieZeit)} disabled={!naechsteFreieZeit}>Übernehmen <kbd class="text-xs opacity-70">Strg+T</kbd></button>
				</div>
				{#if zeitmessung.zeiten.length}
					<ul class="max-h-72 divide-y divide-line overflow-y-auto border-t border-orange-200 bg-surface">
						{#each [...zeitmessung.zeiten].reverse().slice(0, 50) as z (z.zeile)}
							{@const belegt = zugeordnet.get(z.id)}
							<li>
								<button class="flex w-full items-center gap-3 px-4 py-1.5 text-left hover:bg-sunken {belegt ? 'opacity-55' : ''}" onclick={() => zeitUebernehmen(z)}>
									<span class="w-14 font-mono text-xs text-muted">#{z.id}</span>
									<span class="display flex-1 text-xl tabular">{formatZeit(z.zeit)}</span>
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

<Dialog offen={klassenAbschluss !== null} titel="{klassenAbschluss?.klasse.name ?? ''} komplett erfasst" breite="max-w-md" onschliessen={() => (klassenAbschluss = null)}>
	{#if klassenAbschluss}
		<div class="flex items-start gap-3">
			<Flag class="mt-0.5 shrink-0 text-accent" />
			<p class="text-sm">Alle Läufe der {klassenAbschluss.klasse.name} sind erfasst. Die Rangliste ist fertig berechnet.</p>
		</div>
		<ol class="mt-4 divide-y divide-line rounded-lg border border-line text-sm">
			{#each abschlussZeilen.filter((z) => z.platz !== null).slice(0, 3) as z (z.starter.id)}
				<li class="flex items-center gap-3 px-3 py-2">
					<span class="w-6 font-bold tabular">{z.platz}.</span>
					<span class="flex-1">{anzeigeName(z.starter)}</span>
					<span class="font-semibold tabular">{formatZeit(z.gesamt)} s</span>
				</li>
			{:else}
				<li class="px-3 py-2 text-muted">Keine gewerteten Fahrer.</li>
			{/each}
		</ol>
		{#if klassenAbschluss.naechste}
			<p class="mt-3 text-xs text-muted">
				Als Nächstes: Nr. {klassenAbschluss.naechste.starter.startnummer} · {s.klasseVon(klassenAbschluss.naechste.starter)?.name} · {LAUF_NAMEN[klassenAbschluss.naechste.lauf]}
			</p>
		{:else}
			<p class="mt-3 text-xs text-ok">Alle Klassen sind vollständig erfasst.</p>
		{/if}
	{/if}
	{#snippet aktionen()}
		<button class="btn" onclick={() => klassenAbschluss && goto(`/veranstaltung/${s.id}/ergebnisse?klasse=${klassenAbschluss.klasse.id}`)}>
			<Trophy size={16} /> Ergebnis anzeigen
		</button>
		{#if klassenAbschluss?.naechste}
			<!-- svelte-ignore a11y_autofocus -->
			<button class="btn btn-primary" onclick={naechsteKlasse} autofocus>
				Zur nächsten Klasse <ArrowRight size={16} />
			</button>
		{:else}
			<button class="btn btn-primary" onclick={() => goto(`/veranstaltung/${s.id}/abschluss`)}>Drucken & Export <ArrowRight size={16} /></button>
		{/if}
	{/snippet}
</Dialog>
