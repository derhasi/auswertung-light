<script lang="ts">
	import type { LaufEingabe, LaufNr, Regeln } from '$lib/domain/typen';
	import type { WertungsZeile } from '$lib/domain/wertung';
	import { formatPunkte, formatZeit } from '$lib/domain/zahlen';

	interface Props {
		zeilen: readonly WertungsZeile[];
		regeln: Regeln & { fehler1Name: string; fehler2Name: string };
		adressen?: boolean;
		training?: boolean;
		druck?: boolean;
	}

	let { zeilen, regeln, adressen = false, training = true, druck = false }: Props = $props();

	const laeufe = $derived((training ? [0, 1, 2] : [1, 2]) as LaufNr[]);
	const kopf: Record<LaufNr, string> = { 0: 'Training', 1: 'Lauf 1', 2: 'Lauf 2' };

	function strafen(l: LaufEingabe | undefined): string {
		if (!l || l.zeit === null) return '';
		const teile = [];
		if (l.fehler1) teile.push(`${l.fehler1}×${regeln.fehler1Name.slice(0, 1)}`);
		if (l.fehler2) teile.push(`${l.fehler2}×${regeln.fehler2Name.slice(0, 1)}`);
		return teile.length ? `${formatZeit(l.zeit)} + ${teile.join(' ')}` : '';
	}

	const zelle = $derived(druck ? 'px-1.5 py-1' : 'px-3 py-2');
</script>

<table class="w-full {druck ? 'text-[10.5pt]' : 'text-sm'}">
	<thead class="text-left text-xs tracking-wide text-muted uppercase {druck ? 'border-b-2 border-fg' : 'bg-sunken'}">
		<tr>
			<th class="{zelle} w-12 text-right font-semibold">Pl.</th>
			<th class="{zelle} w-12 text-right font-semibold">Nr.</th>
			<th class="{zelle} font-semibold">Fahrer</th>
			{#each laeufe as nr (nr)}<th class="{zelle} text-right font-semibold">{kopf[nr]}</th>{/each}
			<th class="{zelle} text-right font-semibold">Gesamt</th>
			<th class="{zelle} text-right font-semibold" title="Wertungspunkte (Mannschaftswertung)">Pkt.</th>
			<th class="{zelle} text-right font-semibold" title="ADAC-Sportabzeichenpunkte">SA</th>
		</tr>
	</thead>
	<tbody>
		{#each zeilen as z (z.starter.id)}
			<tr class="border-t border-line align-top {druck ? 'break-inside-avoid' : ''} {z.status !== 'gewertet' ? 'text-muted' : ''}">
				<td class="{zelle} text-right font-bold tabular">
					{#if z.status === 'ausser-wertung'}niW{:else if z.platz}{z.platz}.{:else}–{/if}
				</td>
				<td class="{zelle} text-right tabular">{z.starter.startnummer}</td>
				<td class={zelle}>
					<span class="font-semibold text-fg">{z.starter.nachname}, {z.starter.vorname}</span>
					{#if z.rookie}<span class="badge ml-1 {druck ? 'border border-current px-1 py-0' : 'bg-info-soft text-info'}">Rookie</span>{/if}
					<div class="text-xs text-muted">
						{z.starter.verein}{#if adressen && (z.starter.plz || z.starter.ort)} · {z.starter.plz} {z.starter.ort}{/if}
					</div>
				</td>
				{#each laeufe as nr (nr)}
					<td class="{zelle} text-right tabular {nr === 0 ? 'text-muted' : ''}">
						{formatZeit(z.ergebnisse[nr])}
						{#if strafen(z.starter.laeufe[nr])}<div class="text-[10px] text-muted">{strafen(z.starter.laeufe[nr])}</div>{/if}
					</td>
				{/each}
				<td class="{zelle} text-right font-bold tabular">{formatZeit(z.gesamt)}</td>
				<td class="{zelle} text-right tabular">{z.status === 'gewertet' ? formatPunkte(z.punkte) : ''}</td>
				<td class="{zelle} text-right tabular">{z.status === 'gewertet' ? formatPunkte(z.sportabzeichen) : ''}</td>
			</tr>
		{/each}
	</tbody>
</table>
