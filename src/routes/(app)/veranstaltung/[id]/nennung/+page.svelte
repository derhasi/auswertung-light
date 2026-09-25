<script lang="ts">
	import { onMount, tick } from 'svelte';
	import { Trash2, UserPlus, X } from '@lucide/svelte';
	import { repo, type Fahrer } from '$lib/db';
	import type { Starter } from '$lib/domain/typen';
	import FahrerSuche from '$lib/components/FahrerSuche.svelte';
	import { ui } from '$lib/ui/ui-zustand.svelte';

	let { data } = $props();
	const s = $derived(data.store);

	let fahrer = $state<Fahrer[]>([]);
	let suche: FahrerSuche | undefined = $state();
	let nummerFeld: HTMLInputElement | undefined = $state();

	interface Entwurf {
		lizenz: string;
		nachname: string;
		vorname: string;
		verein: string;
		plz: string;
		ort: string;
		rookieJahr: number | null;
		klasseId: number;
		startnummer: number;
		ausserWertung: boolean;
		manuell: boolean;
	}
	let entwurf = $state<Entwurf | null>(null);

	onMount(async () => {
		try {
			fahrer = await (await repo()).fahrerListe();
		} catch (e) {
			ui.fehler(e);
		}
	});

	const leer = $derived(s.klassen.filter((k) => !s.starter.some((st) => st.klasseId === k.id)));
	const gemeldet = $derived(new Set(s.starter.map((st) => st.lizenz).filter(Boolean)));

	function passendeKlasse(klasse: string): number {
		const k = klasse.trim().toLowerCase();
		const treffer = s.klassen.find((kl) => kl.kuerzel.toLowerCase() === k || kl.name.toLowerCase() === k);
		return (treffer ?? s.klassen[0])?.id ?? 0;
	}

	async function vorbereiten(f: Fahrer | null) {
		const klasseId = f ? passendeKlasse(f.klasse) : (s.klassen[0]?.id ?? 0);
		entwurf = {
			lizenz: f?.lizenz ?? '',
			nachname: f?.nachname ?? '',
			vorname: f?.vorname ?? '',
			verein: f?.verein ?? '',
			plz: f?.plz ?? '',
			ort: f?.ort ?? '',
			rookieJahr: f?.rookieJahr ?? null,
			klasseId,
			startnummer: s.naechsteStartnummer(klasseId),
			ausserWertung: false,
			manuell: !f
		};
		await tick();
		nummerFeld?.select();
	}

	async function nennen(e: SubmitEvent) {
		e.preventDefault();
		if (!entwurf) return;
		const { manuell: _m, ...d } = entwurf;
		if (!d.nachname.trim()) return ui.melden('Bitte einen Namen angeben.', 'warnung');
		if (d.lizenz && gemeldet.has(d.lizenz)) {
			const ok = await ui.bestaetigen(`${d.vorname} ${d.nachname} ist bereits gemeldet. Trotzdem ein weiteres Mal nennen (z. B. in einer anderen Klasse)?`, { ja: 'Trotzdem nennen' });
			if (!ok) return;
		}
		try {
			await s.nennen({ ...d, startnummer: Number(d.startnummer), rookieJahr: d.rookieJahr ? Number(d.rookieJahr) : null });
			ui.melden(`Nr. ${d.startnummer} – ${d.vorname} ${d.nachname} gemeldet.`);
			entwurf = null;
			await tick();
			suche?.fokussieren();
		} catch (e) {
			ui.fehler(e);
		}
	}

	async function aendern(st: Starter, daten: Partial<Starter>, feld?: HTMLInputElement | HTMLSelectElement) {
		try {
			await s.startAktualisieren(st.id, daten);
		} catch (e) {
			ui.fehler(e);
			// Anzeige auf den gespeicherten Stand zurücksetzen
			if (feld instanceof HTMLInputElement && feld.type === 'checkbox') feld.checked = st.ausserWertung;
			else if (feld) feld.value = String(feld instanceof HTMLSelectElement ? st.klasseId : st.startnummer);
		}
	}

	async function entfernen(st: Starter) {
		const mitErgebnissen = Object.keys(st.laeufe).length > 0;
		const ok = await ui.bestaetigen(
			`Nr. ${st.startnummer} – ${st.vorname} ${st.nachname} aus der Nennliste entfernen?${mitErgebnissen ? '\n\nAchtung: Die bereits erfassten Läufe werden ebenfalls gelöscht.' : ''}`,
			{ ja: 'Entfernen', gefaehrlich: true }
		);
		if (ok) await s.startLoeschen(st.id).catch((e) => ui.fehler(e));
	}
</script>

<div class="grid gap-6 px-8 py-6">
	<section class="card p-5">
		<div class="flex flex-wrap items-center gap-3">
			<div class="min-w-72 flex-1">
				<FahrerSuche bind:this={suche} {fahrer} {gemeldet} onauswahl={vorbereiten} />
			</div>
			<button class="btn" onclick={() => vorbereiten(null)}><UserPlus size={16} /> Ohne Datenbank nennen</button>
		</div>
		{#if fahrer.length === 0}
			<p class="mt-3 text-sm text-muted">Die Fahrerdatenbank ist leer. <a class="text-accent-strong underline" href="/fahrer">Fahrerliste importieren</a> oder Fahrer ohne Datenbank nennen.</p>
		{/if}

		{#if entwurf}
			<form class="mt-4 rounded-xl border border-accent/40 bg-accent-soft/40 p-4" onsubmit={nennen}>
				<div class="mb-3 flex items-center justify-between">
					<p class="font-semibold">
						{#if entwurf.manuell}Fahrer ohne Datenbank nennen{:else}{entwurf.nachname}, {entwurf.vorname} <span class="font-normal text-muted">· {entwurf.verein} · Lizenz {entwurf.lizenz}</span>{/if}
					</p>
					<button type="button" class="btn btn-ghost btn-icon" onclick={() => (entwurf = null)} aria-label="Abbrechen"><X size={16} /></button>
				</div>
				{#if entwurf.manuell}
					<div class="mb-3 grid grid-cols-2 gap-3 md:grid-cols-4">
						<div><label class="label" for="m-nachname">Nachname</label><input id="m-nachname" class="input" bind:value={entwurf.nachname} required /></div>
						<div><label class="label" for="m-vorname">Vorname</label><input id="m-vorname" class="input" bind:value={entwurf.vorname} /></div>
						<div><label class="label" for="m-verein">Verein</label><input id="m-verein" class="input" bind:value={entwurf.verein} /></div>
						<div><label class="label" for="m-lizenz">Lizenz (optional)</label><input id="m-lizenz" class="input" bind:value={entwurf.lizenz} /></div>
						<div><label class="label" for="m-plz">PLZ</label><input id="m-plz" class="input" bind:value={entwurf.plz} /></div>
						<div><label class="label" for="m-ort">Wohnort</label><input id="m-ort" class="input" bind:value={entwurf.ort} /></div>
						<div><label class="label" for="m-rookie">Rookie-Jahr</label><input id="m-rookie" class="input" type="number" bind:value={entwurf.rookieJahr} /></div>
					</div>
				{/if}
				<div class="flex flex-wrap items-end gap-3">
					<div>
						<label class="label" for="n-klasse">Klasse</label>
						<select
							id="n-klasse"
							class="input w-44"
							bind:value={entwurf.klasseId}
							onchange={() => entwurf && (entwurf.startnummer = s.naechsteStartnummer(entwurf.klasseId))}
						>
							{#each s.klassen as k (k.id)}<option value={k.id}>{k.name}</option>{/each}
						</select>
					</div>
					<div>
						<label class="label" for="n-nummer">Startnummer</label>
						<input id="n-nummer" bind:this={nummerFeld} class="input w-32 font-semibold tabular" type="number" min="1" bind:value={entwurf.startnummer} required />
					</div>
					<label class="flex items-center gap-2 pb-2 text-sm">
						<input type="checkbox" class="size-4 accent-[var(--color-accent)]" bind:checked={entwurf.ausserWertung} /> außer Wertung (niW)
					</label>
					<button class="btn btn-primary ml-auto" type="submit"><UserPlus size={16} /> Nennen <kbd class="text-xs opacity-70">↵</kbd></button>
				</div>
				{#if s.nachStartnummer.has(Number(entwurf.startnummer))}
					<p class="mt-2 text-sm text-danger">Startnummer {entwurf.startnummer} ist bereits vergeben.</p>
				{/if}
			</form>
		{/if}
	</section>

	{#each s.klassen.filter((k) => s.starter.some((st) => st.klasseId === k.id)) as klasse (klasse.id)}
		{@const liste = s.starter.filter((st) => st.klasseId === klasse.id)}
		<section class="card overflow-hidden">
			<header class="flex items-center justify-between border-b border-line px-5 py-3">
				<h2 class="font-semibold">{klasse.name} <span class="ml-1 text-sm font-normal text-muted">{liste.length} Starter</span></h2>
			</header>
			<table class="w-full text-sm">
				<tbody>
					{#each liste as st (st.id)}
						<tr class="border-t border-line first:border-t-0">
							<td class="w-24 py-1.5 pl-5">
								<input
									class="input w-20 py-1 font-semibold tabular"
									type="number"
									min="1"
									value={st.startnummer}
									aria-label="Startnummer"
									onchange={(e) => aendern(st, { startnummer: Number(e.currentTarget.value) }, e.currentTarget)}
								/>
							</td>
							<td class="px-3 py-1.5">
								<span class="font-medium">{st.nachname}, {st.vorname}</span>
								{#if st.rookieJahr !== null && st.rookieJahr === s.jahr}<span class="badge ml-1 bg-info-soft text-info">Rookie</span>{/if}
								{#if !st.lizenz}<span class="badge ml-1 bg-sunken text-muted">ohne Lizenz</span>{/if}
							</td>
							<td class="px-3 py-1.5 text-muted">{st.verein}</td>
							<td class="px-3 py-1.5 text-muted">{st.plz} {st.ort}</td>
							<td class="px-3 py-1.5">
								<select class="input w-36 py-1" value={st.klasseId} aria-label="Klasse" onchange={(e) => aendern(st, { klasseId: Number(e.currentTarget.value) }, e.currentTarget)}>
									{#each s.klassen as k (k.id)}<option value={k.id}>{k.name}</option>{/each}
								</select>
							</td>
							<td class="px-3 py-1.5">
								<label class="flex items-center gap-1.5 text-xs whitespace-nowrap text-muted">
									<input type="checkbox" class="size-4 accent-[var(--color-accent)]" checked={st.ausserWertung} onchange={(e) => aendern(st, { ausserWertung: e.currentTarget.checked }, e.currentTarget)} />
									niW
								</label>
							</td>
							<td class="w-12 pr-3 text-right">
								<button class="btn btn-ghost btn-icon text-muted hover:text-danger" onclick={() => entfernen(st)} aria-label="Nennung entfernen"><Trash2 size={16} /></button>
							</td>
						</tr>
					{/each}
				</tbody>
			</table>
		</section>
	{/each}
	{#if leer.length}
		<p class="text-sm text-muted">Noch ohne Nennungen: {leer.map((k) => k.name).join(', ')}</p>
	{/if}
</div>
