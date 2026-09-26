<script lang="ts">
	import { FileDown, FileUp, Plus, Search, Trash2, TriangleAlert } from '@lucide/svelte';
	import { repo, type Fahrer, type FahrerStart, type FahrerVersion } from '$lib/db';
	import { lizenzGueltig } from '$lib/domain/typen';
	import { dekodiereText, stringifyCsv } from '$lib/domain/csv';
	import { FAHRER_CSV_KOPF, fahrerCsvZeile, formatDatum, parseDatum, parseFahrerCsv, type FahrerDaten } from '$lib/domain/fahrer-import';
	import Seitenkopf from '$lib/components/Seitenkopf.svelte';
	import Dialog from '$lib/ui/Dialog.svelte';
	import { ui } from '$lib/ui/ui-zustand.svelte';
	import { CSV_FILTER, dateiOeffnen, dateiSpeichern } from '$lib/plattform';

	type Sortierung = 'name' | 'lizenz' | 'klasse' | 'verein';

	let fahrer = $state<Fahrer[]>([]);
	let starts = $state(new Map<number, FahrerStart[]>());
	let versionen = $state<FahrerVersion[]>([]);
	let suche = $state('');
	let sortierung = $state<Sortierung>('name');
	let klassenFilter = $state('');
	let bearbeiten = $state<(FahrerDaten & { id?: number; geburtsdatumText: string }) | null>(null);

	async function laden() {
		const r = await repo();
		[fahrer, starts] = await Promise.all([r.fahrerListe(), r.fahrerStarts()]);
	}
	laden().catch((e) => ui.fehler(e));

	const klassen = $derived([...new Set(fahrer.map((f) => f.klasse).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'de', { numeric: true })));

	const gefiltert = $derived.by(() => {
		const begriffe = suche.toLocaleLowerCase('de').split(/\s+/).filter(Boolean);
		const liste = fahrer.filter(
			(f) =>
				(!klassenFilter || f.klasse === klassenFilter) &&
				begriffe.every((b) => `${f.lizenz} ${f.nachname} ${f.vorname} ${f.verein} ${f.ort}`.toLocaleLowerCase('de').includes(b))
		);
		const vergleich: Record<Sortierung, (a: Fahrer, b: Fahrer) => number> = {
			name: (a, b) => a.nachname.localeCompare(b.nachname, 'de') || a.vorname.localeCompare(b.vorname, 'de'),
			lizenz: (a, b) => a.lizenz.localeCompare(b.lizenz, 'de', { numeric: true }),
			klasse: (a, b) => a.klasse.localeCompare(b.klasse, 'de', { numeric: true }) || vergleich.name(a, b),
			verein: (a, b) => a.verein.localeCompare(b.verein, 'de') || vergleich.name(a, b)
		};
		return liste.sort(vergleich[sortierung]);
	});


	async function importieren() {
		try {
			const datei = await dateiOeffnen('Fahrerliste importieren', CSV_FILTER);
			if (!datei) return;
			const { fahrer: liste, fehler } = parseFahrerCsv(dekodiereText(datei.bytes));
			if (!liste.length) {
				ui.melden('In der Datei wurden keine Fahrer gefunden.', 'fehler', fehler);
				return;
			}
			const ergebnis = await (await repo()).fahrerImportieren(liste);
			await laden();
			ui.melden(
				`${liste.length} Fahrer gelesen: ${ergebnis.neu} neu, ${ergebnis.aktualisiert} aktualisiert, ${ergebnis.unveraendert} unverändert.`,
				fehler.length ? 'warnung' : 'erfolg',
				fehler
			);
		} catch (e) {
			ui.fehler(e, 'Import fehlgeschlagen');
		}
	}

	async function exportieren() {
		try {
			const csv = stringifyCsv([FAHRER_CSV_KOPF, ...fahrer.map(fahrerCsvZeile)]);
			const ziel = await dateiSpeichern('fahrer.csv', csv, CSV_FILTER);
			if (ziel) ui.melden(`${fahrer.length} Fahrer exportiert.`);
		} catch (e) {
			ui.fehler(e, 'Export fehlgeschlagen');
		}
	}

	function oeffnen(f?: Fahrer) {
		versionen = [];
		if (f) repo().then((r) => r.fahrerVersionen(f.id)).then((v) => (versionen = v)).catch((e) => ui.fehler(e));
		bearbeiten = f
			? { ...f, geburtsdatumText: formatDatum(f.geburtsdatum) }
			: { lizenz: '', klasse: '', nachname: '', vorname: '', rookieJahr: null, plz: '', ort: '', verein: '', geburtsdatum: '', alteLizenz: '', geburtsdatumText: '' };
	}

	async function speichern(e: SubmitEvent) {
		e.preventDefault();
		if (!bearbeiten) return;
		const { geburtsdatumText, ...daten } = bearbeiten;
		const lizenz = daten.lizenz.trim();
		if (lizenz && !lizenzGueltig(lizenz)) {
			ui.melden('Die Lizenz darf nur Buchstaben, Ziffern sowie - / _ enthalten.', 'warnung');
			return;
		}
		const vergeben = fahrer.find((f) => lizenz && f.lizenz === lizenz && f.id !== daten.id);
		if (vergeben) {
			ui.melden(`Die Lizenz ${lizenz} ist bereits an ${vergeben.vorname} ${vergeben.nachname} vergeben. Lizenzen müssen eindeutig sein.`, 'fehler');
			return;
		}
		try {
			await (await repo()).fahrerSpeichern({
				...daten,
				lizenz,
				geburtsdatum: parseDatum(geburtsdatumText),
				rookieJahr: daten.rookieJahr ? Number(daten.rookieJahr) : null
			});
			bearbeiten = null;
			await laden();
			ui.melden('Fahrer gespeichert.');
		} catch (e) {
			ui.fehler(e);
		}
	}

	async function loeschen() {
		if (!bearbeiten?.id) return;
		const name = `${bearbeiten.vorname} ${bearbeiten.nachname}`;
		if (!(await ui.bestaetigen(`${name} aus der Fahrerdatenbank löschen?\nBereits gemeldete Starts bleiben erhalten.`, { ja: 'Löschen', gefaehrlich: true }))) return;
		await (await repo()).fahrerLoeschen([bearbeiten.id]);
		bearbeiten = null;
		await laden();
	}
</script>

<Seitenkopf titel="Fahrerdatenbank" untertitel="{fahrer.length} Fahrer · wird für alle Veranstaltungen verwendet">
	{#snippet aktionen()}
		<button class="btn" onclick={importieren}><FileUp size={16} /> ZP-Fahrerliste importieren</button>
		<button class="btn" onclick={exportieren} disabled={!fahrer.length}><FileDown size={16} /> Exportieren</button>
		<button class="btn btn-primary" onclick={() => oeffnen()}><Plus size={16} /> Fahrer anlegen</button>
	{/snippet}
</Seitenkopf>

<div class="px-8 pb-10">
	<div class="mb-4 flex flex-wrap items-center gap-3">
		<div class="relative w-80 max-w-full">
			<Search size={16} class="pointer-events-none absolute top-1/2 left-3 -translate-y-1/2 text-muted" />
			<input class="input pl-9" type="search" placeholder="Name, Lizenz, Verein oder Ort suchen" bind:value={suche} />
		</div>
		<select class="input w-auto" bind:value={klassenFilter} aria-label="Klasse filtern">
			<option value="">Alle Klassen</option>
			{#each klassen as k (k)}<option value={k}>{k}</option>{/each}
		</select>
		<label class="ml-auto flex items-center gap-2 text-sm text-muted">
			Sortieren nach
			<select class="input w-auto" bind:value={sortierung}>
				<option value="name">Name</option>
				<option value="lizenz">Lizenz</option>
				<option value="klasse">Klasse</option>
				<option value="verein">Verein</option>
			</select>
		</label>
	</div>

	{#if fahrer.length === 0}
		<div class="card p-8 text-center text-sm text-muted">
			<p>Noch keine Fahrer vorhanden.</p>
			<p class="mt-1">
				Importiere die CSV-Fahrerliste von zugspitzpokal.de (Spalten: ID, Klasse, Nachname, Vorname, Rookie, PLZ, Wohnort, Verein, Geburtsdatum).
			</p>
		</div>
	{:else}
		<div class="card overflow-hidden">
			<table class="w-full text-sm">
				<thead class="bg-sunken text-left text-xs tracking-wide text-muted uppercase">
					<tr>
						<th class="px-4 py-2.5 font-semibold">Lizenz</th>
						<th class="px-4 py-2.5 font-semibold">Name</th>
						<th class="px-4 py-2.5 font-semibold">Klasse</th>
						<th class="px-4 py-2.5 font-semibold">Verein</th>
						<th class="px-4 py-2.5 font-semibold">Wohnort</th>
						<th class="px-4 py-2.5 font-semibold">Geburtsdatum</th>
						<th class="px-4 py-2.5 font-semibold">Starts</th>
					</tr>
				</thead>
				<tbody>
					{#each gefiltert.slice(0, 500) as f (f.id)}
						<tr class="cursor-pointer border-t border-line hover:bg-sunken/60" onclick={() => oeffnen(f)}>
							<td class="px-4 py-2 font-mono text-xs">
								{f.lizenz}
								{#if !lizenzGueltig(f.lizenz)}<span class="badge ml-1 bg-warn-soft text-warn" title="Lizenz enthält unzulässige Zeichen (erlaubt: Buchstaben, Ziffern, - / _)"><TriangleAlert size={11} /></span>{/if}
							</td>
							<td class="px-4 py-2 font-medium">
								{f.nachname}, {f.vorname}
								{#if f.rookieJahr}<span class="badge ml-1 bg-info-soft text-info">Rookie {f.rookieJahr}</span>{/if}
							</td>
							<td class="px-4 py-2">{f.klasse}</td>
							<td class="px-4 py-2">{f.verein}</td>
							<td class="px-4 py-2 text-muted">{f.plz} {f.ort}</td>
							<td class="px-4 py-2 text-muted tabular">{formatDatum(f.geburtsdatum)}</td>
							<td class="px-4 py-2">
								{#if starts.get(f.id)?.length}<span class="badge bg-ok-soft text-ok">{starts.get(f.id)?.length}</span>{/if}
							</td>
						</tr>
					{/each}
				</tbody>
			</table>
			{#if gefiltert.length > 500}
				<p class="border-t border-line px-4 py-2 text-xs text-muted">Es werden die ersten 500 von {gefiltert.length} Treffern angezeigt – bitte die Suche verfeinern.</p>
			{:else if gefiltert.length === 0}
				<p class="border-t border-line px-4 py-6 text-center text-sm text-muted">Keine Treffer.</p>
			{/if}
		</div>
	{/if}
</div>

<Dialog offen={bearbeiten !== null} titel={bearbeiten?.id ? 'Fahrer bearbeiten' : 'Fahrer anlegen'} onschliessen={() => (bearbeiten = null)}>
	{#if bearbeiten}
		<form id="fahrer" class="grid grid-cols-2 gap-4" onsubmit={speichern}>
			<div>
				<label class="label" for="f-lizenz">Lizenz</label>
				<input id="f-lizenz" class="input font-mono" bind:value={bearbeiten.lizenz} required />
			</div>
			<div>
				<label class="label" for="f-klasse">Klasse</label>
				<input id="f-klasse" class="input" bind:value={bearbeiten.klasse} placeholder="K1" list="klassen-liste" />
				<datalist id="klassen-liste">{#each klassen as k (k)}<option value={k}></option>{/each}</datalist>
			</div>
			<div>
				<label class="label" for="f-nachname">Nachname</label>
				<input id="f-nachname" class="input" bind:value={bearbeiten.nachname} required />
			</div>
			<div>
				<label class="label" for="f-vorname">Vorname</label>
				<input id="f-vorname" class="input" bind:value={bearbeiten.vorname} />
			</div>
			<div class="col-span-2">
				<label class="label" for="f-verein">Verein</label>
				<input id="f-verein" class="input" bind:value={bearbeiten.verein} />
				<p class="mt-1 text-xs text-muted">Für die Mannschaftswertung bei allen Fahrern gleich schreiben.</p>
			</div>
			<div>
				<label class="label" for="f-plz">PLZ</label>
				<input id="f-plz" class="input" bind:value={bearbeiten.plz} />
			</div>
			<div>
				<label class="label" for="f-ort">Wohnort</label>
				<input id="f-ort" class="input" bind:value={bearbeiten.ort} />
			</div>
			<div>
				<label class="label" for="f-geb">Geburtsdatum</label>
				<input id="f-geb" class="input" bind:value={bearbeiten.geburtsdatumText} placeholder="TT.MM.JJJJ" />
			</div>
			<div>
				<label class="label" for="f-rookie">Rookie-Jahr</label>
				<input id="f-rookie" class="input" type="number" min="1990" max="2100" bind:value={bearbeiten.rookieJahr} />
			</div>
			<div class="col-span-2">
				<label class="label" for="f-alt">Alte Lizenz-Nr.</label>
				<input id="f-alt" class="input" bind:value={bearbeiten.alteLizenz} />
			</div>
			{#if bearbeiten.id && starts.get(bearbeiten.id)?.length}
				<div class="col-span-2">
					<p class="label">Gestartet bei</p>
					<ul class="divide-y divide-line rounded-lg border border-line text-sm">
						{#each starts.get(bearbeiten.id) ?? [] as s, i (i)}
							<li class="flex justify-between px-3 py-1.5">
								<a class="hover:text-accent-strong" href="/veranstaltung/{s.veranstaltungId}">{formatDatum(s.datum)} · {s.veranstaltung}</a>
								<span class="text-muted">
									{s.klasse} · Nr. {s.startnummer}
									{#if s.fahrerVersionId && versionen.length > 1}· Version {versionen.length - versionen.findIndex((v) => v.id === s.fahrerVersionId)}{/if}
								</span>
							</li>
						{/each}
					</ul>
				</div>
			{/if}
			{#if versionen.length > 1}
				<div class="col-span-2">
					<p class="label">Frühere Stände ({versionen.length} Versionen)</p>
					<ul class="divide-y divide-line rounded-lg border border-line text-xs">
						{#each versionen as v, i (v.id)}
							<li class="px-3 py-1.5">
								<div class="flex justify-between">
									<span class="font-semibold">Version {versionen.length - i}{i === 0 ? ' (aktuell)' : ''}</span>
									<span class="text-muted">{new Date(v.erstelltAm).toLocaleString('de-DE', { dateStyle: 'short', timeStyle: 'short' })} · {v.anlass}</span>
								</div>
								<div class="text-muted">
									{v.nachname}, {v.vorname} · {v.verein} · {v.plz} {v.ort} · Lizenz {v.lizenz}
								</div>
							</li>
						{/each}
					</ul>
				</div>
			{/if}
		</form>
	{/if}
	{#snippet aktionen()}
		{#if bearbeiten?.id}
			<button class="btn btn-ghost mr-auto text-danger" onclick={loeschen}><Trash2 size={16} /> Löschen</button>
		{/if}
		<button class="btn" onclick={() => (bearbeiten = null)}>Abbrechen</button>
		<button class="btn btn-primary" type="submit" form="fahrer">Speichern</button>
	{/snippet}
</Dialog>
