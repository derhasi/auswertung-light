<script lang="ts">
	import { goto } from '$app/navigation';
	import { ArrowLeft, Printer } from '@lucide/svelte';
	import DruckKopf from '$lib/components/DruckKopf.svelte';
	import ErgebnisListe from '$lib/components/ErgebnisListe.svelte';
	import MannschaftsListe from '$lib/components/MannschaftsListe.svelte';
	import { formatDatum } from '$lib/domain/fahrer-import';
	import { anzeigeName } from '$lib/domain/typen';
	import { formatZeit } from '$lib/domain/zahlen';
	import { drucken } from '$lib/plattform';
	import { VERSION } from '$lib/version';

	let { data } = $props();
	const s = $derived(data.store);

	const dokumente = {
		start: 'Startliste',
		ergebnis: 'Ergebnisliste',
		mannschaft: 'Mannschaftswertung',
		urkunden: 'Urkunden'
	} as const;
	type Dok = keyof typeof dokumente;
	const dok = $derived((data.dok in dokumente ? data.dok : 'ergebnis') as Dok);

	const klassen = $derived(
		s.klassenWertungen.filter((k) => k.zeilen.length > 0 && (data.klasse === 'alle' || String(k.klasse.id) === data.klasse))
	);
	const urkunden = $derived(
		klassen.flatMap(({ klasse, zeilen }) =>
			zeilen.filter((z) => z.status === 'gewertet' && z.platz !== null && z.platz <= s.v.urkundenPlaetze).map((z) => ({ klasse, z }))
		)
	);

	function parameter(aenderung: Record<string, string>) {
		const url = new URL(window.location.href);
		for (const [k, v] of Object.entries(aenderung)) url.searchParams.set(k, v);
		goto(url.pathname + url.search, { replaceState: true });
	}

	const gedruckt = new Date().toLocaleString('de-DE', { dateStyle: 'short', timeStyle: 'short' });
</script>

<svelte:head>
	<title>{dokumente[dok]} – {s.v.name}</title>
	{@html `<style>@page { size: A4 ${dok === 'ergebnis' && data.training ? 'landscape' : 'portrait'}; margin: 12mm; }</style>`}
</svelte:head>

<div class="sticky top-0 z-10 flex flex-wrap items-center gap-3 border-b border-line bg-surface px-6 py-3 print:hidden">
	<a class="btn btn-ghost" href="/veranstaltung/{s.id}/abschluss"><ArrowLeft size={16} /> Zurück</a>
	<select class="input w-auto" value={dok} onchange={(e) => parameter({ dok: e.currentTarget.value })} aria-label="Dokument">
		{#each Object.entries(dokumente) as [wert, titel] (wert)}<option value={wert}>{titel}</option>{/each}
	</select>
	{#if dok !== 'mannschaft'}
		<select class="input w-auto" value={data.klasse} onchange={(e) => parameter({ klasse: e.currentTarget.value })} aria-label="Klasse">
			<option value="alle">Alle Klassen</option>
			{#each s.klassen as k (k.id)}<option value={String(k.id)}>{k.name}</option>{/each}
		</select>
	{/if}
	{#if dok === 'ergebnis'}
		<label class="flex items-center gap-2 text-sm"><input type="checkbox" checked={data.training} onchange={(e) => parameter({ training: e.currentTarget.checked ? '1' : '0' })} /> Training</label>
		<label class="flex items-center gap-2 text-sm"><input type="checkbox" checked={data.adressen} onchange={(e) => parameter({ adressen: e.currentTarget.checked ? '1' : '0' })} /> Adressen</label>
	{/if}
	<button class="btn btn-primary ml-auto" onclick={drucken}><Printer size={16} /> Drucken / PDF</button>
</div>

<div class="druckbereich mx-auto my-6 max-w-[1100px] bg-surface p-10 text-fg shadow-sm print:m-0 print:max-w-none print:p-0 print:shadow-none">
	{#if dok === 'mannschaft'}
		<DruckKopf v={s.v} titel="Mannschaftswertung" />
		<MannschaftsListe zeilen={s.mannschaft} anzahl={s.v.mannschaftAnzahl} druck />
	{:else if dok === 'urkunden'}
		{#each urkunden as { klasse, z }, i (z.starter.id)}
			<section class="urkunde flex min-h-[260mm] flex-col items-center text-center {i > 0 ? 'break-before-page' : ''}">
				<div class="flex w-full items-start justify-between">
					<div class="h-28 w-40">{#if s.v.logoLinks}<img src={s.v.logoLinks} alt="" class="max-h-28 max-w-40 object-contain" />{/if}</div>
					<div class="h-28 w-40 text-right">{#if s.v.logoRechts}<img src={s.v.logoRechts} alt="" class="ml-auto max-h-28 max-w-40 object-contain" />{/if}</div>
				</div>
				<h1 class="mt-10 text-[44pt] font-black tracking-[0.2em] uppercase">Urkunde</h1>
				<p class="mt-6 text-[14pt]">{s.v.name}</p>
				<p class="text-[12pt] text-muted">{formatDatum(s.v.datum)}{s.v.ort ? ` in ${s.v.ort}` : ''}</p>
				<p class="mt-16 text-[28pt] font-bold">{z.starter.vorname} {z.starter.nachname}</p>
				<p class="text-[14pt] text-muted">{z.starter.verein}</p>
				<p class="mt-14 text-[16pt]">belegte in der {klasse.name} den</p>
				<p class="mt-2 text-[60pt] leading-none font-black">{z.platz}. Platz</p>
				<p class="mt-6 text-[12pt] text-muted">Gesamtzeit {formatZeit(z.gesamt)} s</p>
				<div class="mt-auto grid w-full grid-cols-2 gap-24 px-10 pb-6 text-[10pt]">
					<div class="border-t border-fg pt-1">Veranstalter</div>
					<div class="border-t border-fg pt-1">Fahrtleiter</div>
				</div>
			</section>
		{:else}
			<p class="text-center text-muted">Noch keine platzierten Fahrer (Urkunden für Platz 1–{s.v.urkundenPlaetze}).</p>
		{/each}
	{:else}
		{#each klassen as { klasse, zeilen }, i (klasse.id)}
			<section class={i > 0 ? 'break-before-page print:pt-0' : ''} class:mt-12={i > 0}>
				<DruckKopf v={s.v} titel={dokumente[dok]} untertitel={klasse.name} />
				{#if dok === 'ergebnis'}
					<ErgebnisListe {zeilen} regeln={s.v} adressen={data.adressen} training={data.training} druck />
				{:else}
					<table class="w-full text-[10.5pt]">
						<thead class="border-b-2 border-fg text-left text-xs text-muted uppercase">
							<tr>
								<th class="w-14 px-1.5 py-1 text-right">Nr.</th>
								<th class="px-1.5 py-1">Fahrer</th>
								<th class="px-1.5 py-1">Verein</th>
								<th class="px-1.5 py-1">Lizenz</th>
								<th class="w-24 px-1.5 py-1 text-center">Training</th>
								<th class="w-24 px-1.5 py-1 text-center">Lauf 1</th>
								<th class="w-24 px-1.5 py-1 text-center">Lauf 2</th>
							</tr>
						</thead>
						<tbody>
							{#each [...zeilen].sort((a, b) => a.starter.startnummer - b.starter.startnummer) as z (z.starter.id)}
								<tr class="break-inside-avoid border-t border-line">
									<td class="px-1.5 py-2 text-right font-bold tabular">{z.starter.startnummer}</td>
									<td class="px-1.5 py-2 font-semibold">
										{anzeigeName(z.starter)}
										{#if z.rookie}<span class="text-xs font-normal">(Rookie)</span>{/if}
										{#if z.starter.ausserWertung}<span class="text-xs font-normal">(niW)</span>{/if}
									</td>
									<td class="px-1.5 py-2">{z.starter.verein}</td>
									<td class="px-1.5 py-2 font-mono text-xs">{z.starter.lizenz}</td>
									<td class="border-l border-line"></td>
									<td class="border-l border-line"></td>
									<td class="border-l border-line"></td>
								</tr>
							{/each}
						</tbody>
					</table>
				{/if}
				<p class="mt-3 text-right text-[8pt] text-muted">
					Strafen: {s.v.fehler1Name} {s.v.strafe1} s · {s.v.fehler2Name} {s.v.strafe2} s · Stand {gedruckt} · Auswertung Light {VERSION}
				</p>
			</section>
		{:else}
			<p class="text-center text-muted">Keine Fahrer gemeldet.</p>
		{/each}
	{/if}
</div>

<style>
	@media print {
		:global(html),
		:global(body) {
			background: white;
		}
		.urkunde {
			min-height: 270mm;
		}
	}
</style>
