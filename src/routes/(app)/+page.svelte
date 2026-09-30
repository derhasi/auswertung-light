<script lang="ts">
	import { goto } from '$app/navigation';
	import { ArrowRight, FileUp, Users } from '@lucide/svelte';
	import { repo, type VeranstaltungsUebersicht } from '$lib/db';
	import { formatDatum } from '$lib/domain/fahrer-import';
	import Seitenkopf from '$lib/components/Seitenkopf.svelte';
	import { ui } from '$lib/ui/ui-zustand.svelte';
	import { dateiOeffnen, JSON_FILTER } from '$lib/plattform';
	import { dekodiereText } from '$lib/domain/csv';

	let liste = $state<VeranstaltungsUebersicht[] | null>(null);
	let fahrerAnzahl = $state(0);
	let neu = $state({ name: '', datum: new Date().toISOString().slice(0, 10), ort: '', vorlage: 'letzte' as string });

	async function laden() {
		try {
			const r = await repo();
			[liste, fahrerAnzahl] = await Promise.all([r.veranstaltungsListe(), r.fahrerListe().then((f) => f.length)]);
		} catch (e) {
			ui.fehler(e, 'Datenbank konnte nicht geöffnet werden');
		}
	}
	laden();

	async function anlegen(e: SubmitEvent) {
		e.preventDefault();
		if (!neu.name.trim()) return;
		try {
			const vorlage = neu.vorlage === 'letzte' ? undefined : neu.vorlage === 'keine' ? null : Number(neu.vorlage);
			const id = await (await repo()).veranstaltungAnlegen({ name: neu.name.trim(), datum: neu.datum, ort: neu.ort.trim() }, vorlage);
			goto(`/veranstaltung/${id}/nennung`);
		} catch (e) {
			ui.fehler(e);
		}
	}

	const heute = new Date().toISOString().slice(0, 10);
	const aktuell = $derived(liste?.[0]);
	const weitere = $derived(liste?.slice(1) ?? []);

	function tag(datum: string) {
		return datum.slice(8, 10);
	}
	function monat(datum: string) {
		const d = new Date(`${datum}T12:00:00`);
		return Number.isNaN(d.getTime()) ? '' : d.toLocaleDateString('de-DE', { month: 'short', year: 'numeric' }).replace('.', '');
	}
	function zustand(datum: string) {
		if (datum === heute) return { text: 'Heute', klasse: 'bg-accent text-on-accent' };
		if (datum > heute) return { text: 'Geplant', klasse: 'bg-sunken text-fg' };
		return { text: 'Vergangen', klasse: 'bg-ok-soft text-ok' };
	}

	async function importieren() {
		try {
			const datei = await dateiOeffnen('Veranstaltung importieren', JSON_FILTER);
			if (!datei) return;
			const id = await (await repo()).veranstaltungImportieren(JSON.parse(dekodiereText(datei.bytes)));
			ui.melden('Veranstaltung importiert.');
			goto(`/veranstaltung/${id}`);
		} catch (e) {
			ui.fehler(e, 'Import fehlgeschlagen');
		}
	}
</script>

<Seitenkopf titel="Veranstaltungen" untertitel="Wähle eine Veranstaltung oder lege eine neue an.">
	{#snippet aktionen()}
		<button class="btn" onclick={importieren}><FileUp size={16} /> Sicherung einlesen (JSON)</button>
	{/snippet}
</Seitenkopf>

<div class="grid gap-6 px-7 py-6 lg:grid-cols-[minmax(0,1fr)_360px]">
	<section class="flex min-w-0 flex-col gap-3.5">
		{#if liste === null}
			<p class="text-muted">Lade …</p>
		{:else if !aktuell}
			<div class="card p-7">
				<h2 class="section-title text-[28px]">Willkommen bei Auswertung Light</h2>
				<p class="mt-1 text-muted">In drei Schritten zur ersten Auswertung:</p>
				<ol class="mt-5 flex flex-col gap-5">
					<li class="flex gap-4">
						<span class="display flex size-12 shrink-0 items-center justify-center rounded-[10px] bg-ink text-[28px] text-on-ink">1</span>
						<div>
							<p class="text-lg font-bold">Fahrerdatenbank füllen</p>
							<p class="text-muted">
								Die Fahrerliste des Zugspitzpokals (CSV) importieren oder Fahrer von Hand anlegen.
								{#if fahrerAnzahl}<span class="badge bg-ok-soft text-ok">{fahrerAnzahl} Fahrer vorhanden</span>{/if}
							</p>
							<a class="btn btn-sm mt-2" href="/fahrer"><Users size={14} /> Zur Fahrerdatenbank</a>
						</div>
					</li>
					<li class="flex gap-4">
						<span class="display flex size-12 shrink-0 items-center justify-center rounded-[10px] bg-accent text-[28px] text-on-accent">2</span>
						<div>
							<p class="text-lg font-bold">Veranstaltung anlegen und Fahrer nennen</p>
							<p class="text-muted">Rechts Bezeichnung und Datum eintragen. Klassen, Strafsekunden und Logos lassen sich je Veranstaltung einstellen.</p>
						</div>
					</li>
					<li class="flex gap-4">
						<span class="display flex size-12 shrink-0 items-center justify-center rounded-[10px] bg-ink text-[28px] text-on-ink">3</span>
						<div>
							<p class="text-lg font-bold">Läufe erfassen, Ergebnisse drucken, ZP-Export speichern</p>
							<p class="text-muted">Ergebnisliste und Mannschaftswertung werden live berechnet.</p>
						</div>
					</li>
				</ol>
			</div>
		{:else}
			{@const z = zustand(aktuell.datum)}
			<h2 class="section-title">Aktuell</h2>
			<a href="/veranstaltung/{aktuell.id}" class="card group flex overflow-hidden hover:border-accent">
				<div class="flex w-36 shrink-0 flex-col items-center justify-center bg-ink py-4 text-on-ink">
					<span class="display text-[72px] leading-none tabular">{tag(aktuell.datum)}</span>
					<span class="eyebrow text-ink-muted">{monat(aktuell.datum)}</span>
				</div>
				<div class="flex min-w-0 flex-1 flex-col justify-center gap-2.5 px-7 py-5">
					<div class="flex flex-wrap gap-2">
						<span class="badge px-2.5 py-1 text-[13px] {z.klasse}">{z.text}</span>
						<span class="badge bg-sunken px-2.5 py-1 text-[13px] text-fg">{aktuell.starter} Starter · {aktuell.klassen} Klassen</span>
					</div>
					<span class="display text-[46px] leading-none group-hover:text-accent">{aktuell.name}</span>
					{#if aktuell.ort}<span class="text-base text-muted">{aktuell.ort}</span>{/if}
				</div>
				<span class="display flex items-center gap-2 px-7 text-[22px] text-accent">Öffnen <ArrowRight size={22} /></span>
			</a>

			{#if weitere.length}
				<h2 class="section-title mt-3">Weitere Veranstaltungen</h2>
				<div class="grid gap-3.5 [grid-template-columns:repeat(auto-fill,minmax(260px,1fr))]">
					{#each weitere as v (v.id)}
						{@const zv = zustand(v.datum)}
						<a href="/veranstaltung/{v.id}" class="card group flex flex-col gap-2 p-4.5 hover:border-accent">
							<div class="flex items-center justify-between">
								<span class="display text-[22px] tabular">{formatDatum(v.datum)}</span>
								<span class="badge {zv.klasse}">{zv.text}</span>
							</div>
							<span class="text-xl leading-tight font-bold group-hover:text-accent">{v.name}</span>
							{#if v.ort}<span class="text-[15px] text-muted">{v.ort}</span>{/if}
							<span class="text-sm font-semibold text-muted">{v.starter} Starter · {v.klassen} Klassen</span>
						</a>
					{/each}
				</div>
			{/if}
		{/if}
	</section>

	<aside class="card flex flex-col gap-4 self-start p-5.5">
		<h2 class="section-title text-[26px]">Neue Veranstaltung</h2>
		<form id="neu" class="flex flex-col gap-4" onsubmit={anlegen}>
			<div>
				<label class="label" for="neu-name">Bezeichnung</label>
				<input id="neu-name" class="input" bind:value={neu.name} placeholder="z. B. 3. Lauf Zugspitzpokal" required />
			</div>
			<div class="grid grid-cols-2 gap-3">
				<div>
					<label class="label" for="neu-datum">Datum</label>
					<input id="neu-datum" class="input" type="date" bind:value={neu.datum} required />
				</div>
				<div>
					<label class="label" for="neu-ort">Ort</label>
					<input id="neu-ort" class="input" bind:value={neu.ort} />
				</div>
			</div>
			<div>
				<label class="label" for="neu-vorlage">Einstellungen übernehmen</label>
				<select id="neu-vorlage" class="input" bind:value={neu.vorlage}>
					<option value="letzte">von der zuletzt angelegten Veranstaltung</option>
					{#each liste ?? [] as v (v.id)}<option value={String(v.id)}>von „{v.name}“ ({formatDatum(v.datum)})</option>{/each}
					<option value="keine">keine – Standard (Klasse 1–6, 2 s / 10 s)</option>
				</select>
				<p class="mt-1.5 text-sm text-muted">Übernommen werden Klassen, Strafsekunden, Logos, Ausrichter und Zeitmessung – keine Fahrer.</p>
			</div>
			<button class="btn btn-primary btn-lg" type="submit">Anlegen</button>
		</form>
	</aside>
</div>
