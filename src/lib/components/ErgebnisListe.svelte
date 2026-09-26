<script lang="ts">
	import { Pencil } from '@lucide/svelte';
	import type { LaufEingabe, LaufNr, Regeln, Starter } from '$lib/domain/typen';
	import { STATUS_KURZ, type WertungsZeile } from '$lib/domain/wertung';
	import { formatPunkte, formatZeit } from '$lib/domain/zahlen';

	interface Props {
		zeilen: readonly WertungsZeile[];
		regeln: Regeln & { fehler1Name: string; fehler2Name: string };
		adressen?: boolean;
		training?: boolean;
		druck?: boolean;
		/** Blendet eine Schaltfläche zum Bearbeiten der Läufe ein. */
		onbearbeiten?: (s: Starter) => void;
	}

	let { zeilen, regeln, adressen = false, training = true, druck = false, onbearbeiten }: Props = $props();

	const laeufe = $derived((training ? [0, 1, 2] : [1, 2]) as LaufNr[]);
	const kopf: Record<LaufNr, string> = { 0: 'Training', 1: 'Lauf 1', 2: 'Lauf 2' };

	function strafen(l: LaufEingabe | undefined): string {
		if (!l || (l.status ?? 'ok') !== 'ok') return l?.kommentar ?? '';
		if (l.zeit === null) return '';
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
			{#if onbearbeiten && !druck}<th class="{zelle} w-10"><span class="sr-only">Bearbeiten</span></th>{/if}
		</tr>
	</thead>
	<tbody>
		{#each zeilen as z (z.starter.id)}
			<tr class="border-t border-line align-top {druck ? 'break-inside-avoid' : ''} {z.status !== 'gewertet' ? 'text-muted' : ''}">
				<td class="{zelle} text-right font-bold tabular">
					{#if z.platz}{z.platz}.{:else}{STATUS_KURZ[z.status]}{/if}
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
					{@const l = z.starter.laeufe[nr]}
					<td class="{zelle} text-right tabular {nr === 0 ? 'text-muted' : ''}" title={l?.kommentar ?? undefined}>
						{#if (l?.status ?? 'ok') !== 'ok'}<span class="font-bold text-danger">{l?.status?.toUpperCase()}</span>{:else}{formatZeit(z.ergebnisse[nr])}{/if}
						{#if strafen(z.starter.laeufe[nr])}<div class="text-[10px] text-muted">{strafen(z.starter.laeufe[nr])}</div>{/if}
					</td>
				{/each}
				<td class="{zelle} text-right font-bold tabular">{formatZeit(z.gesamt)}</td>
				<td class="{zelle} text-right tabular">{z.status === 'gewertet' ? formatPunkte(z.punkte) : ''}</td>
				<td class="{zelle} text-right tabular">{z.status === 'gewertet' ? formatPunkte(z.sportabzeichen) : ''}</td>
				{#if onbearbeiten && !druck}
					<td class="{zelle} text-right">
						<button class="btn btn-ghost btn-icon -my-1 text-muted hover:text-fg" onclick={() => onbearbeiten(z.starter)} aria-label="Läufe von Nr. {z.starter.startnummer} bearbeiten" title="Läufe bearbeiten">
							<Pencil size={15} />
						</button>
					</td>
				{/if}
			</tr>
		{/each}
	</tbody>
</table>
