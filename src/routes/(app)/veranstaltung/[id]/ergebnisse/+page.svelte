<script lang="ts">
	import { Printer } from '@lucide/svelte';
	import ErgebnisListe from '$lib/components/ErgebnisListe.svelte';

	let { data } = $props();
	const s = $derived(data.store);

	let auswahl = $state<number | 'alle'>('alle');
	let adressen = $state(false);
	let training = $state(true);

	const sichtbar = $derived(
		s.klassenWertungen.filter((k) => k.zeilen.length > 0 && (auswahl === 'alle' || k.klasse.id === auswahl))
	);
</script>

<div class="px-8 py-6">
	<div class="mb-5 flex flex-wrap items-center gap-2">
		<button class="chip" aria-pressed={auswahl === 'alle'} onclick={() => (auswahl = 'alle')}>Alle Klassen</button>
		{#each s.klassenWertungen as { klasse, zeilen } (klasse.id)}
			{#if zeilen.length}
				<button class="chip" aria-pressed={auswahl === klasse.id} onclick={() => (auswahl = klasse.id)}>
					{klasse.name} <span class="text-xs opacity-70">{zeilen.length}</span>
				</button>
			{/if}
		{/each}
		<div class="ml-auto flex items-center gap-4 text-sm">
			<label class="flex items-center gap-2"><input type="checkbox" class="size-4 accent-[var(--color-accent)]" bind:checked={training} /> Training</label>
			<label class="flex items-center gap-2"><input type="checkbox" class="size-4 accent-[var(--color-accent)]" bind:checked={adressen} /> Adressen</label>
			<a class="btn" href="/druck/{s.id}?dok=ergebnis&klasse={auswahl}&adressen={adressen ? 1 : 0}&training={training ? 1 : 0}">
				<Printer size={16} /> Drucken
			</a>
		</div>
	</div>

	{#if sichtbar.length === 0}
		<div class="card p-8 text-center text-sm text-muted">Noch keine Fahrer gemeldet.</div>
	{/if}
	<div class="flex flex-col gap-6">
		{#each sichtbar as { klasse, zeilen } (klasse.id)}
			<section class="card overflow-hidden">
				<header class="flex items-center justify-between border-b border-line px-5 py-3">
					<h2 class="font-semibold">{klasse.name}</h2>
					<span class="text-xs text-muted">
						{zeilen.filter((z) => z.status === 'gewertet').length} gewertet
						{#if zeilen.some((z) => z.status === 'unvollstaendig')}
							· <span class="text-warn">{zeilen.filter((z) => z.status === 'unvollstaendig').length} unvollständig</span>
						{/if}
					</span>
				</header>
				<ErgebnisListe {zeilen} regeln={s.v} {adressen} {training} />
			</section>
		{/each}
	</div>
</div>
