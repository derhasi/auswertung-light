<script lang="ts">
	import { ArrowRight, ClipboardList } from '@lucide/svelte';
	import { anzeigeName, LAUF_NAMEN, laufErfasst, type LaufNr } from '$lib/domain/typen';
	import { naechsterStart } from '$lib/domain/reihenfolge';
	import { formatDatum } from '$lib/domain/fahrer-import';
	import { formatZeit } from '$lib/domain/zahlen';

	let { data } = $props();
	const s = $derived(data.store);
	const basis = $derived(`/veranstaltung/${s.id}`);

	const LAEUFE: { nr: LaufNr; titel: string }[] = [
		{ nr: 0, titel: 'Training' },
		{ nr: 1, titel: 'Lauf 1' },
		{ nr: 2, titel: 'Lauf 2' }
	];

	type Zeilen = (typeof s.klassenWertungen)[number]['zeilen'];
	function erfasst(zeilen: Zeilen, nr: LaufNr) {
		return zeilen.filter((z) => laufErfasst(z.starter.laeufe[nr])).length;
	}

	const laeufe = $derived({
		gesamt: s.starter.length * 3,
		erfasst: s.starter.reduce((a, st) => a + ([0, 1, 2] as LaufNr[]).filter((nr) => laufErfasst(st.laeufe[nr])).length, 0)
	});
	const klassenMitFahrern = $derived(s.klassenWertungen.filter((k) => k.zeilen.length));
	const naechster = $derived(naechsterStart(s.reihenfolge));
	const naechsteKlasse = $derived(naechster ? s.klassen.find((k) => k.id === naechster.starter.klasseId) : undefined);
	const offenInKlasse = $derived(
		naechster ? s.reihenfolge.filter((p) => p.starter.klasseId === naechster.starter.klasseId && p.lauf === naechster.lauf && !laufErfasst(p.starter.laeufe[p.lauf])).length : 0
	);
</script>

<header class="flex flex-wrap items-center gap-6 border-b border-line bg-surface px-7 py-5">
	<div class="min-w-0">
		<h1 class="display text-[44px] leading-none">{s.v.name}</h1>
		<p class="mt-1.5 text-base text-muted">
			{formatDatum(s.v.datum)}{s.v.ort ? ` · ${s.v.ort}` : ''}{s.v.ausrichter ? ` · Ausrichter ${s.v.ausrichter}` : ''}
		</p>
	</div>
	<div class="ml-auto flex gap-3">
		<div class="card-ink flex w-32 flex-col justify-center px-4 py-2.5">
			<span class="eyebrow text-xs text-ink-muted">Starter</span>
			<span class="display text-[40px] leading-none tabular">{s.starter.length}</span>
		</div>
		<div class="card-ink flex w-32 flex-col justify-center px-4 py-2.5">
			<span class="eyebrow text-xs text-ink-muted">Klassen</span>
			<span class="display text-[40px] leading-none tabular">{klassenMitFahrern.length}</span>
		</div>
		<div class="flex w-40 flex-col justify-center rounded-[14px] bg-accent px-4 py-2.5 text-on-accent">
			<span class="eyebrow text-xs">Läufe erfasst</span>
			<span class="display text-[40px] leading-none tabular">{laeufe.erfasst}/{laeufe.gesamt}</span>
		</div>
	</div>
</header>

<div class="grid gap-6 px-7 py-6 lg:grid-cols-[1fr_340px]">
	<section class="card self-start overflow-hidden">
		<h2 class="section-title border-b border-line px-5 py-4 text-2xl">Stand der Klassen</h2>
		{#if s.starter.length === 0}
			<div class="p-10 text-center text-muted">
				<p>Noch keine Fahrer gemeldet.</p>
				<a class="btn btn-primary mt-4" href="{basis}/nennung"><ClipboardList size={16} /> Zur Nennung</a>
			</div>
		{:else}
			<table class="w-full">
				<thead class="border-b-2 border-fg text-left text-[13px] tracking-wider text-muted uppercase">
					<tr>
						<th class="px-5 py-3 font-bold">Klasse</th>
						<th class="px-3 py-3 text-right font-bold">Starter</th>
						{#each LAEUFE as l (l.nr)}<th class="px-3 py-3 font-bold">{l.titel}</th>{/each}
						<th class="px-5 py-3 font-bold">Führung</th>
					</tr>
				</thead>
				<tbody>
					{#each klassenMitFahrern as { klasse, zeilen } (klasse.id)}
						<tr class="border-t border-line">
							<td class="display px-5 py-4 text-[26px]">{klasse.name}</td>
							<td class="display px-3 py-4 text-right text-[26px] tabular">{zeilen.length}</td>
							{#each LAEUFE as l (l.nr)}
								{@const n = erfasst(zeilen, l.nr)}
								{@const voll = n === zeilen.length}
								<td class="px-3 py-4">
									<div class="h-2 rounded bg-sunken">
										<div class="h-2 rounded {voll ? 'bg-fg' : 'bg-accent'}" style:width="{(n / zeilen.length) * 100}%"></div>
									</div>
									<div class="mt-1.5 text-[13px] font-semibold tabular {voll ? 'text-muted' : 'text-accent'}">{n}/{zeilen.length}</div>
								</td>
							{/each}
							<td class="px-5 py-4">
								{#if zeilen[0]?.platz === 1}
									<div class="font-bold">{zeilen[0].starter.vorname} {zeilen[0].starter.nachname}</div>
									<div class="text-[13px] text-muted">{zeilen[0].starter.verein}{zeilen[0].gesamt != null ? ` · ${formatZeit(zeilen[0].gesamt)} s` : ''}</div>
								{:else}<span class="text-muted">–</span>{/if}
							</td>
						</tr>
					{/each}
				</tbody>
			</table>
		{/if}
	</section>

	<aside class="flex flex-col gap-4">
		<div class="card-ink flex flex-col gap-2.5 p-5">
			{#if naechster && naechsteKlasse}
				<span class="eyebrow text-ink-muted">Weiter mit</span>
				<span class="display text-[34px] leading-none">{naechsteKlasse.name} · {LAUF_NAMEN[naechster.lauf]}</span>
				<span class="text-[15px] text-ink-muted">
					{offenInKlasse === 1 ? '1 Fahrer offen' : `${offenInKlasse} Fahrer offen`} · als Nächstes Nr. {naechster.starter.startnummer}
					{anzeigeName(naechster.starter)}
				</span>
				<a href="{basis}/erfassung" class="btn btn-primary btn-lg mt-1.5">Zur Erfassung</a>
			{:else if s.starter.length}
				<span class="eyebrow text-ink-muted">Erfassung</span>
				<span class="display text-[34px] leading-none">Alle Läufe erfasst</span>
				<a href="{basis}/ergebnisse" class="btn btn-primary btn-lg mt-1.5">Zu den Ergebnissen</a>
			{:else}
				<span class="eyebrow text-ink-muted">Als Erstes</span>
				<span class="display text-[34px] leading-none">Fahrer nennen</span>
				<a href="{basis}/nennung" class="btn btn-primary btn-lg mt-1.5">Zur Nennung</a>
			{/if}
		</div>
		<nav aria-label="Schnellzugriff" class="card flex flex-col py-2">
			<span class="eyebrow px-5 py-2 text-muted">Schnellzugriff</span>
			{#each [{ href: '/nennung', t: 'Fahrer nennen' }, { href: '/ergebnisse', t: 'Ergebnisse ansehen' }, { href: '/mannschaft', t: 'Mannschaftswertung' }, { href: '/abschluss', t: 'Drucken & Export' }] as l (l.href)}
				<a href="{basis}{l.href}" class="flex items-center justify-between border-t border-line px-5 py-3 font-semibold hover:bg-sunken">
					{l.t} <ArrowRight size={16} class="text-accent" />
				</a>
			{/each}
		</nav>
		<div class="card p-5 text-[15px]">
			<p class="eyebrow text-muted">Regeln dieser Veranstaltung</p>
			<ul class="mt-2 space-y-1">
				<li>{s.v.fehler1Name}: {s.v.strafe1} s je Fehler</li>
				<li>{s.v.fehler2Name}: {s.v.strafe2} s je Fehler</li>
				<li>Mannschaft: beste {s.v.mannschaftAnzahl} Ergebnisse je Verein</li>
			</ul>
		</div>
	</aside>
</div>
