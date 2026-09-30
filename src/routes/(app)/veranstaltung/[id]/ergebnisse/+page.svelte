<script lang="ts">
	import { page } from '$app/state';
	import { Printer } from '@lucide/svelte';
	import ErgebnisListe from '$lib/components/ErgebnisListe.svelte';
	import Seitenkopf from '$lib/components/Seitenkopf.svelte';
	import { formatPunkte, formatZeit } from '$lib/domain/zahlen';
	import LaufBearbeiten from '$lib/components/LaufBearbeiten.svelte';
	import type { Starter } from '$lib/domain/typen';

	let { data } = $props();
	const s = $derived(data.store);

	// Klasse aus der Adresse übernehmen (z. B. aus dem Zwischenschritt der Erfassung)
	const klasseParam = Number(page.url.searchParams.get('klasse'));
	let auswahl = $state<number | 'alle'>(Number.isFinite(klasseParam) && klasseParam > 0 ? klasseParam : 'alle');
	let adressen = $state(false);
	let training = $state(true);
	let bearbeiten = $state<Starter | null>(null);

	const sichtbar = $derived(
		s.klassenWertungen.filter((k) => k.zeilen.length > 0 && (auswahl === 'alle' || k.klasse.id === auswahl))
	);
	/** Podest nur bei einer einzelnen Klasse. */
	const podest = $derived(auswahl === 'alle' || sichtbar.length !== 1 ? [] : sichtbar[0].zeilen.filter((z) => z.platz !== null && z.platz <= 3));
</script>

<Seitenkopf titel="Ergebnisse">
	{#snippet mitte()}
		<nav aria-label="Klasse wählen" class="flex flex-wrap gap-1.5">
			<button class="chip" aria-pressed={auswahl === 'alle'} onclick={() => (auswahl = 'alle')}>Alle Klassen</button>
			{#each s.klassenWertungen as { klasse, zeilen } (klasse.id)}
				{#if zeilen.length}
					<button class="chip" aria-pressed={auswahl === klasse.id} onclick={() => (auswahl = klasse.id)}>{klasse.name}</button>
				{/if}
			{/each}
		</nav>
	{/snippet}
	{#snippet aktionen()}
		<div class="flex items-center gap-4 text-[15px] font-semibold">
			<label class="flex items-center gap-2"><input type="checkbox" class="size-5 accent-[var(--color-accent)]" bind:checked={training} /> Training</label>
			<label class="flex items-center gap-2"><input type="checkbox" class="size-5 accent-[var(--color-accent)]" bind:checked={adressen} /> Adressen</label>
			<a class="btn btn-ink" href="/druck/{s.id}?dok=ergebnis&klasse={auswahl}&adressen={adressen ? 1 : 0}&training={training ? 1 : 0}">
				<Printer size={16} /> Ergebnisliste drucken
			</a>
		</div>
	{/snippet}
</Seitenkopf>

<div class="flex flex-col gap-5 px-7 py-6">
	{#if podest.length}
		<div class="grid gap-3.5 md:grid-cols-3">
			{#each podest as z (z.starter.id)}
				<div class="card-ink flex items-center gap-4 px-5 py-4">
					<span class="display text-[72px] leading-none {z.platz === 1 ? 'text-accent' : ''}">{z.platz}</span>
					<div class="flex min-w-0 flex-1 flex-col">
						<span class="display truncate text-[28px] leading-tight">{z.starter.vorname} {z.starter.nachname}</span>
						<span class="truncate text-sm text-ink-muted">{z.starter.verein}</span>
					</div>
					<div class="flex flex-col items-end">
						<span class="display text-[34px] leading-none tabular">{formatZeit(z.gesamt)}</span>
						<span class="text-[13px] text-ink-muted tabular">{formatPunkte(z.punkte)} Pkt.</span>
					</div>
				</div>
			{/each}
		</div>
	{/if}

	{#if sichtbar.length === 0}
		<div class="card p-10 text-center text-muted">Noch keine Fahrer gemeldet.</div>
	{/if}
	{#each sichtbar as { klasse, zeilen } (klasse.id)}
		<section class="card overflow-hidden">
			<header class="flex flex-wrap items-baseline justify-between gap-2 border-b border-line px-5 py-3.5">
				<h2 class="section-title text-2xl">{klasse.name} · {zeilen.length} Starter</h2>
				<span class="text-sm text-muted">
					{zeilen.filter((z) => z.status === 'gewertet').length} gewertet
					{#if zeilen.some((z) => z.status === 'unvollstaendig')}
						· <span class="text-warn">{zeilen.filter((z) => z.status === 'unvollstaendig').length} unvollständig</span>
					{/if}
					{#if zeilen.some((z) => z.status === 'nicht-gestartet' || z.status === 'disqualifiziert')}
						· <span class="text-danger">{zeilen.filter((z) => z.status === 'nicht-gestartet' || z.status === 'disqualifiziert').length} DNS/DSQ</span>
					{/if}
				</span>
			</header>
			<ErgebnisListe {zeilen} regeln={s.v} {adressen} {training} onbearbeiten={(st) => (bearbeiten = st)} />
		</section>
	{/each}
	<p class="text-sm text-muted">Gesamt = Lauf 1 + Lauf 2 · bei Gleichstand entscheidet der bessere Einzellauf · Pkt. = Wertungspunkte für die Mannschaft · SA = Sportabzeichen-Punkte</p>
</div>

<LaufBearbeiten store={s} starter={bearbeiten} onschliessen={() => (bearbeiten = null)} />
