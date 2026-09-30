<script lang="ts">
	import type { MannschaftsZeile } from '$lib/domain/mannschaft';
	import { formatPunkte, formatZeit } from '$lib/domain/zahlen';

	let { zeilen, anzahl, druck = false }: { zeilen: readonly MannschaftsZeile[]; anzahl: number; druck?: boolean } = $props();
	const zelle = $derived(druck ? 'px-1.5 py-1' : 'px-3 py-2');
</script>

<table class="w-full {druck ? 'text-[10pt]' : 'text-sm'}">
	<thead class="text-left text-xs tracking-wide text-muted uppercase {druck ? 'border-b-2 border-fg' : 'bg-sunken'}">
		<tr>
			<th class="{zelle} w-12 text-right font-semibold">Pl.</th>
			<th class="{zelle} font-semibold">Verein</th>
			<th class="{zelle} text-right font-semibold">Summe</th>
			<th class="{zelle} font-semibold">Wertende Ergebnisse (beste {anzahl})</th>
		</tr>
	</thead>
	<tbody>
		{#each zeilen as z (z.verein)}
			<tr class="border-t border-line align-top {druck ? 'break-inside-avoid' : ''}">
				<td class="{zelle} text-right font-bold tabular">{z.platz}.</td>
				<td class="{zelle} font-semibold">{z.verein}</td>
				<td class="{zelle} text-right text-base font-bold tabular">{formatZeit(z.summe)}</td>
				<td class={zelle}>
					<div class="flex flex-wrap gap-1.5">
						{#each z.ergebnisse as e, i (i)}
							<span class="inline-flex items-baseline gap-1.5 rounded-md border border-line px-2 py-0.5 text-xs">
								<span class="font-semibold tabular">{formatPunkte(e.punkte)}</span>
								<span>{e.fahrer.join(' & ')}</span>
								<span class="text-muted">{e.klasse}</span>
							</span>
						{/each}
					</div>
				</td>
			</tr>
		{/each}
	</tbody>
</table>
