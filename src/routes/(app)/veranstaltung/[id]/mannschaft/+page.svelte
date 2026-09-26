<script lang="ts">
	import { Printer } from '@lucide/svelte';
	import MannschaftsListe from '$lib/components/MannschaftsListe.svelte';

	let { data } = $props();
	const s = $derived(data.store);
	const ohne = $derived(s.klassen.filter((k) => !k.inMannschaft));
</script>

<div class="px-8 py-6">
	<div class="mb-5 flex flex-wrap items-center gap-3">
		<p class="text-sm text-muted">
			Je Verein zählen die besten {s.v.mannschaftAnzahl} Punktergebnisse aus allen Klassen.
			{#if ohne.length}Nicht berücksichtigt: {ohne.map((k) => k.name).join(', ')}.{/if}
			Vereinsnamen werden ohne Beachtung von Groß-/Kleinschreibung zusammengefasst.
		</p>
		<a class="btn ml-auto" href="/druck/{s.id}?dok=mannschaft"><Printer size={16} /> Drucken</a>
	</div>
	{#if s.mannschaft.length === 0}
		<div class="card p-8 text-center text-sm text-muted">Noch keine gewerteten Ergebnisse.</div>
	{:else}
		<div class="card overflow-hidden">
			<MannschaftsListe zeilen={s.mannschaft} anzahl={s.v.mannschaftAnzahl} />
		</div>
	{/if}
</div>
