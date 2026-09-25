<script lang="ts">
	import { ArrowRight, CircleCheck, ClipboardList, Timer } from '@lucide/svelte';
	import { LAUF_KURZ, WERTUNGSLAEUFE, type LaufNr } from '$lib/domain/typen';

	let { data } = $props();
	const s = $derived(data.store);

	function erfasst(zeilen: typeof s.klassenWertungen[number]['zeilen'], nr: LaufNr) {
		return zeilen.filter((z) => z.ergebnisse[nr] !== null).length;
	}
	const gesamt = $derived({
		starter: s.starter.length,
		fertig: s.klassenWertungen.reduce((a, k) => a + k.zeilen.filter((z) => z.gesamt !== null).length, 0)
	});
</script>

<div class="grid gap-6 px-8 py-6 lg:grid-cols-[1fr_320px]">
	<section class="card overflow-hidden">
		<h2 class="border-b border-line px-5 py-3 font-semibold">Stand der Klassen</h2>
		{#if s.starter.length === 0}
			<div class="p-8 text-center text-sm text-muted">
				<p>Noch keine Fahrer gemeldet.</p>
				<a class="btn btn-primary mt-3" href="/veranstaltung/{s.id}/nennung"><ClipboardList size={16} /> Zur Nennung</a>
			</div>
		{:else}
			<table class="w-full text-sm">
				<thead class="bg-sunken text-left text-xs tracking-wide text-muted uppercase">
					<tr>
						<th class="px-5 py-2 font-semibold">Klasse</th>
						<th class="px-3 py-2 text-right font-semibold">Starter</th>
						<th class="px-3 py-2 text-right font-semibold">Training</th>
						{#each WERTUNGSLAEUFE as nr (nr)}<th class="px-3 py-2 text-right font-semibold">{LAUF_KURZ[nr]}</th>{/each}
						<th class="px-5 py-2 font-semibold">Führung</th>
					</tr>
				</thead>
				<tbody>
					{#each s.klassenWertungen as { klasse, zeilen } (klasse.id)}
						{#if zeilen.length}
							<tr class="border-t border-line">
								<td class="px-5 py-2.5 font-medium">{klasse.name}</td>
								<td class="px-3 py-2.5 text-right tabular">{zeilen.length}</td>
								{#each [0, ...WERTUNGSLAEUFE] as nr (nr)}
									{@const n = erfasst(zeilen, nr as LaufNr)}
									<td class="px-3 py-2.5 text-right tabular {n === zeilen.length ? 'text-ok' : 'text-muted'}">
										{#if n === zeilen.length}<CircleCheck size={14} class="mr-1 inline" />{/if}{n}/{zeilen.length}
									</td>
								{/each}
								<td class="px-5 py-2.5 text-muted">
									{#if zeilen[0]?.platz === 1}{zeilen[0].starter.nachname}, {zeilen[0].starter.vorname}{:else}–{/if}
								</td>
							</tr>
						{/if}
					{/each}
				</tbody>
			</table>
		{/if}
	</section>

	<aside class="flex flex-col gap-4">
		<div class="card p-5">
			<p class="text-xs font-semibold tracking-wide text-muted uppercase">Fortschritt</p>
			<p class="mt-2 text-3xl font-bold tabular">{gesamt.fertig}<span class="text-lg text-muted"> / {gesamt.starter}</span></p>
			<p class="text-sm text-muted">Fahrer mit beiden Wertungsläufen</p>
			<div class="mt-3 h-2 overflow-hidden rounded-full bg-sunken">
				<div class="h-full bg-accent transition-all" style:width="{gesamt.starter ? (gesamt.fertig / gesamt.starter) * 100 : 0}%"></div>
			</div>
		</div>
		<a class="card flex items-center gap-3 p-4 hover:shadow-md" href="/veranstaltung/{s.id}/erfassung">
			<Timer class="text-accent" />
			<div class="flex-1">
				<p class="font-medium">Läufe erfassen</p>
				<p class="text-xs text-muted">Startnummer, Fehler und Zeit eingeben</p>
			</div>
			<ArrowRight size={16} class="text-muted" />
		</a>
		<div class="card p-4 text-sm">
			<p class="font-medium">Regeln dieser Veranstaltung</p>
			<ul class="mt-2 space-y-1 text-muted">
				<li>{s.v.fehler1Name}: {s.v.strafe1} s je Fehler</li>
				<li>{s.v.fehler2Name}: {s.v.strafe2} s je Fehler</li>
				<li>Mannschaft: beste {s.v.mannschaftAnzahl} Ergebnisse je Verein</li>
			</ul>
		</div>
	</aside>
</div>
