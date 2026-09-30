<script lang="ts">
	import { Printer } from '@lucide/svelte';
	import Seitenkopf from '$lib/components/Seitenkopf.svelte';
	import { formatPunkte, formatZeit } from '$lib/domain/zahlen';

	let { data } = $props();
	const s = $derived(data.store);
	const ohne = $derived(s.klassen.filter((k) => !k.inMannschaft));
	const mit = $derived(s.klassen.filter((k) => k.inMannschaft));
	const hoechste = $derived(Math.max(...s.mannschaft.map((z) => z.summe), 0));
</script>

<Seitenkopf titel="Mannschaftswertung" untertitel="Die {s.v.mannschaftAnzahl} besten Ergebnisse je Verein · {mit.map((k) => k.name).join(', ') || 'keine Klassen'}">
	{#snippet aktionen()}
		<a class="btn btn-ink" href="/druck/{s.id}?dok=mannschaft"><Printer size={16} /> Mannschaftswertung drucken</a>
	{/snippet}
</Seitenkopf>

<div class="grid gap-6 px-7 py-6 lg:grid-cols-[minmax(0,1fr)_340px]">
	<section class="flex min-w-0 flex-col gap-3">
		{#if s.mannschaft.length === 0}
			<div class="card p-10 text-center text-muted">Noch keine gewerteten Ergebnisse.</div>
		{/if}
		{#each s.mannschaft as z (z.verein)}
			<div class="card flex overflow-hidden">
				<div class="display flex w-24 shrink-0 items-center justify-center text-[56px] text-on-ink tabular {z.platz === 1 ? 'bg-accent text-on-accent' : 'bg-ink'}">{z.platz}</div>
				<div class="flex min-w-0 flex-1 flex-col gap-2.5 px-5 py-3.5">
					<div class="flex flex-wrap items-baseline gap-x-4">
						<span class="display text-[30px] leading-none">{z.verein}</span>
						<span class="text-sm text-muted">{z.ergebnisse.length} {z.ergebnisse.length === 1 ? 'Ergebnis' : 'Ergebnisse'} gewertet</span>
						<span class="display ml-auto text-[36px] leading-none tabular">{formatZeit(z.summe)}</span>
					</div>
					<div class="h-2 rounded bg-sunken">
						<div class="h-2 rounded {z.platz === 1 ? 'bg-accent' : 'bg-fg'}" style:width="{hoechste ? (z.summe / hoechste) * 100 : 0}%"></div>
					</div>
					<div class="flex flex-wrap gap-1.5">
						{#each z.ergebnisse as e, i (i)}
							<span class="rounded bg-sunken px-2 py-0.5 text-[13px] font-semibold">
								{e.fahrer.join(' & ')} <span class="text-muted">{e.klasse}</span> · <span class="tabular">{formatPunkte(e.punkte)}</span>
							</span>
						{/each}
					</div>
				</div>
			</div>
		{/each}
	</section>

	<aside class="flex flex-col gap-3.5">
		<div class="card-ink flex flex-col gap-3 p-5">
			<span class="section-title text-2xl">So wird gewertet</span>
			<ul class="flex flex-col gap-2.5 text-[15px] leading-snug text-ink-muted">
				<li>Je Klasse gibt es Punkte nach Platz: (N − Platz) × 10 ÷ N + 1.</li>
				<li>Pro Verein zählen die {s.v.mannschaftAnzahl} besten Ergebnisse aus allen Klassen mit Mannschaftswertung.</li>
				<li>Vereinsnamen werden ohne Beachtung von Groß-/Kleinschreibung zusammengefasst.</li>
			</ul>
		</div>
		<div class="card flex flex-col gap-2 p-5">
			<span class="eyebrow text-muted">Gewertete Klassen</span>
			<div class="flex flex-wrap gap-1.5">
				{#each mit as k (k.id)}<span class="rounded bg-accent px-2.5 py-1 text-sm font-bold text-on-accent">{k.name}</span>{/each}
				{#each ohne as k (k.id)}<span class="rounded bg-sunken px-2.5 py-1 text-sm font-bold text-muted line-through">{k.name}</span>{/each}
			</div>
			<a class="font-semibold text-accent-strong hover:underline" href="/veranstaltung/{s.id}/einstellungen">In den Einstellungen ändern</a>
		</div>
	</aside>
</div>
