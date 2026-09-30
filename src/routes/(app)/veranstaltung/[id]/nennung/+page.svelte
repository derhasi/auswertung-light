<script lang="ts">
	import { onMount, tick } from 'svelte';
	import { FileUp, Trash2, UserPlus, X } from '@lucide/svelte';
	import { repo, type Fahrer } from '$lib/db';
	import { lizenzGueltig, type Starter } from '$lib/domain/typen';
	import FahrerSuche from '$lib/components/FahrerSuche.svelte';
	import NennungImport from '$lib/components/NennungImport.svelte';
	import Seitenkopf from '$lib/components/Seitenkopf.svelte';
	import { ui } from '$lib/ui/ui-zustand.svelte';

	let { data } = $props();
	const s = $derived(data.store);

	let fahrer = $state<Fahrer[]>([]);
	let suche: FahrerSuche | undefined = $state();
	let nummerFeld: HTMLInputElement | undefined = $state();
	let importOffen = $state(false);

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
		/** Datensatz aus der Fahrerdatenbank (bei manueller Eingabe leer). */
		fahrer: Fahrer | null;
		/** Manuell erfasste Fahrer in die Datenbank übernehmen. */
		inDatenbank: boolean;
	}
	let entwurf = $state<Entwurf | null>(null);

	async function fahrerLaden() {
		try {
			fahrer = await (await repo()).fahrerListe();
		} catch (e) {
			ui.fehler(e);
		}
	}
	onMount(fahrerLaden);

	const mitNennung = $derived(s.klassen.filter((k) => s.starter.some((st) => st.klasseId === k.id)));
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
			manuell: !f,
			fahrer: f,
			inDatenbank: true
		};
		await tick();
		nummerFeld?.select();
	}

	async function nennen(e: SubmitEvent) {
		e.preventDefault();
		if (!entwurf) return;
		const { manuell, fahrer: dbFahrer, inDatenbank, ...d } = entwurf;
		d.lizenz = d.lizenz.trim();
		if (!d.nachname.trim()) return ui.melden('Bitte einen Namen angeben.', 'warnung');
		if (manuell && d.lizenz && !lizenzGueltig(d.lizenz)) return ui.melden('Die Lizenz darf nur Buchstaben, Ziffern sowie - / _ enthalten.', 'warnung');
		if (d.lizenz && gemeldet.has(d.lizenz)) {
			const ok = await ui.bestaetigen(`${d.vorname} ${d.nachname} ist bereits gemeldet. Trotzdem ein weiteres Mal nennen (z. B. in einer anderen Klasse)?`, { ja: 'Trotzdem nennen' });
			if (!ok) return;
		}
		try {
			const nennung = { klasseId: d.klasseId, startnummer: Number(d.startnummer), ausserWertung: d.ausserWertung };
			const rookieJahr = d.rookieJahr ? Number(d.rookieJahr) : null;
			if (dbFahrer) {
				// Mit der aktuellen Version der Fahrerdaten verknüpfen
				const ref = await (await repo()).fahrerSpeichern({ ...dbFahrer, id: dbFahrer.id });
				await s.nennenMitFahrer(dbFahrer, ref, nennung);
			} else if (manuell && inDatenbank) {
				const daten = { ...d, rookieJahr, klasse: s.klasseVon(d)?.kuerzel ?? '', geburtsdatum: '', alteLizenz: '' };
				const ref = await (await repo()).fahrerSpeichern(daten, 'Nennung');
				await s.nennenMitFahrer(daten, ref, nennung);
				await fahrerLaden();
			} else {
				await s.nennen({ ...d, ...nennung, rookieJahr });
			}
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

<Seitenkopf titel="Nennung" untertitel="{s.starter.length} Fahrer in {mitNennung.length} {mitNennung.length === 1 ? 'Klasse' : 'Klassen'}">
	{#snippet aktionen()}
		<button class="btn btn-ink" onclick={() => (importOffen = true)}><FileUp size={16} /> Nennliste importieren (CSV/Excel)</button>
	{/snippet}
</Seitenkopf>

<div class="grid items-start gap-6 px-7 py-6 xl:grid-cols-[420px_minmax(0,1fr)]">
	<section class="card overflow-hidden xl:sticky xl:top-6">
		<h2 class="section-title bg-ink px-5 py-3.5 text-2xl text-on-ink">Fahrer nennen</h2>
		<div class="flex flex-col gap-4 p-5">
			<FahrerSuche bind:this={suche} {fahrer} {gemeldet} onauswahl={vorbereiten} />
			{#if !entwurf}
				<button class="btn" onclick={() => vorbereiten(null)}><UserPlus size={16} /> Neuen Fahrer nennen</button>
			{/if}
			{#if fahrer.length === 0}
				<p class="text-sm text-muted">Die Fahrerdatenbank ist leer. <a class="font-semibold text-accent-strong underline" href="/fahrer">Fahrerliste importieren</a>, eine Nennliste importieren oder neue Fahrer direkt nennen.</p>
			{/if}

			{#if entwurf}
				<form class="flex flex-col gap-4 border-t border-line pt-4" onsubmit={nennen}>
					<div class="flex items-start justify-between gap-2">
						<div class="min-w-0">
							{#if entwurf.manuell}
								<p class="display text-[26px] leading-none">Neuer Fahrer</p>
							{:else}
								<p class="display truncate text-[26px] leading-none">{entwurf.vorname} {entwurf.nachname}</p>
								<p class="mt-1 text-sm text-muted">{entwurf.verein} · Lizenz {entwurf.lizenz}</p>
							{/if}
						</div>
						<button type="button" class="btn btn-ghost btn-icon" onclick={() => (entwurf = null)} aria-label="Abbrechen"><X size={18} /></button>
					</div>
					{#if entwurf.manuell}
						<div class="grid grid-cols-2 gap-3">
							<div><label class="label" for="m-nachname">Nachname</label><input id="m-nachname" class="input" bind:value={entwurf.nachname} required /></div>
							<div><label class="label" for="m-vorname">Vorname</label><input id="m-vorname" class="input" bind:value={entwurf.vorname} /></div>
							<div><label class="label" for="m-verein">Verein</label><input id="m-verein" class="input" bind:value={entwurf.verein} /></div>
							<div><label class="label" for="m-lizenz">Lizenz (optional)</label><input id="m-lizenz" class="input" bind:value={entwurf.lizenz} /></div>
							<div><label class="label" for="m-plz">PLZ</label><input id="m-plz" class="input" bind:value={entwurf.plz} /></div>
							<div><label class="label" for="m-ort">Wohnort</label><input id="m-ort" class="input" bind:value={entwurf.ort} /></div>
							<div><label class="label" for="m-rookie">Rookie-Jahr</label><input id="m-rookie" class="input" type="number" bind:value={entwurf.rookieJahr} /></div>
							<label class="flex items-center gap-2 self-end pb-3 text-[15px] font-semibold">
								<input type="checkbox" class="size-5 accent-[var(--color-accent)]" bind:checked={entwurf.inDatenbank} /> In Datenbank
							</label>
						</div>
					{/if}
					<div>
						<label class="label" for="n-klasse">Klasse</label>
						<select
							id="n-klasse"
							class="input"
							bind:value={entwurf.klasseId}
							onchange={() => entwurf && (entwurf.startnummer = s.naechsteStartnummer(entwurf.klasseId))}
						>
							{#each s.klassen as k (k.id)}<option value={k.id}>{k.name}</option>{/each}
						</select>
					</div>
					<div class="flex items-center gap-4">
						<div class="flex w-32 shrink-0 flex-col items-center rounded-[10px] bg-ink py-2 text-on-ink">
							<label class="eyebrow text-xs text-ink-muted" for="n-nummer">Startnr.</label>
							<input id="n-nummer" bind:this={nummerFeld} class="display w-24 bg-transparent text-center text-[52px] leading-none tabular outline-none" type="number" min="1" bind:value={entwurf.startnummer} required />
						</div>
						<div class="flex flex-col gap-2">
							{#if s.nachStartnummer.has(Number(entwurf.startnummer))}
								<span class="text-sm font-semibold text-danger">Startnummer {entwurf.startnummer} ist bereits vergeben.</span>
							{:else}
								<span class="text-sm text-muted">Vorschlag: nächste freie Nummer der Klasse</span>
							{/if}
							<label class="flex items-center gap-2 text-[15px] font-semibold">
								<input type="checkbox" class="size-5 accent-[var(--color-accent)]" bind:checked={entwurf.ausserWertung} /> außer Wertung (niW)
							</label>
						</div>
					</div>
					<div class="flex gap-2.5">
						<button type="button" class="btn flex-1" onclick={() => (entwurf = null)}>Abbrechen</button>
						<button class="btn btn-primary btn-lg flex-[2]" type="submit">Nennen ↵</button>
					</div>
				</form>
			{/if}
		</div>
	</section>

	<div class="flex min-w-0 flex-col gap-3.5">
		{#if mitNennung.length === 0}
			<div class="card p-10 text-center text-muted">Noch keine Nennungen. Links einen Fahrer suchen oder eine Nennliste importieren.</div>
		{/if}
		{#each mitNennung as klasse (klasse.id)}
			{@const liste = s.starter.filter((st) => st.klasseId === klasse.id)}
			<section class="card overflow-hidden">
				<header class="flex items-center justify-between border-b border-line px-5 py-3">
					<h2 class="section-title">{klasse.name}</h2>
					<span class="text-sm font-semibold text-muted">{liste.length} Fahrer</span>
				</header>
				<table class="w-full text-[15px]">
					<tbody>
						{#each liste as st (st.id)}
							<tr class="border-t border-line first:border-t-0">
								<td class="w-24 py-2 pl-4">
									<input
										class="display w-18 rounded-lg border border-transparent bg-sunken px-2 py-1 text-center text-[22px] tabular hover:border-line focus:border-accent focus:outline-none"
										type="number"
										min="1"
										value={st.startnummer}
										aria-label="Startnummer von {st.vorname} {st.nachname}"
										onchange={(e) => aendern(st, { startnummer: Number(e.currentTarget.value) }, e.currentTarget)}
									/>
								</td>
								<td class="px-3 py-2">
									<span class="font-bold">{st.vorname} {st.nachname}</span>
									{#if st.rookieJahr !== null && st.rookieJahr === s.jahr}<span class="badge ml-1 bg-accent-soft text-accent-strong">Rookie</span>{/if}
									{#if !st.lizenz}<span class="badge ml-1 bg-sunken text-muted">ohne Lizenz</span>{/if}
								</td>
								<td class="px-3 py-2 text-muted">{st.verein}</td>
								<td class="px-3 py-2 text-muted">{st.plz} {st.ort}</td>
								<td class="px-3 py-2">
									<select class="input w-36 py-1.5" value={st.klasseId} aria-label="Klasse" onchange={(e) => aendern(st, { klasseId: Number(e.currentTarget.value) }, e.currentTarget)}>
										{#each s.klassen as k (k.id)}<option value={k.id}>{k.name}</option>{/each}
									</select>
								</td>
								<td class="px-3 py-2">
									<label class="flex items-center gap-1.5 text-sm font-semibold whitespace-nowrap text-muted">
										<input type="checkbox" class="size-5 accent-[var(--color-accent)]" checked={st.ausserWertung} onchange={(e) => aendern(st, { ausserWertung: e.currentTarget.checked }, e.currentTarget)} />
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
</div>

<NennungImport store={s} offen={importOffen} onschliessen={() => ((importOffen = false), fahrerLaden())} />
