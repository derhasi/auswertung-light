<script lang="ts">
	import { FileUp, GitMerge, TriangleAlert, UserPlus, Users } from '@lucide/svelte';
	import { repo, type Fahrer } from '$lib/db';
	import { FELD_NAMEN, formatDatum, leseNennungsDatei, type NennungsZeile } from '$lib/domain/fahrer-import';
	import { importPlanen, mitBestandVergleichen, startnummernVergeben, zusammengefuehrt, type ImportPosten, type KonfliktLoesung } from '$lib/domain/nennung-import';
	import { dateiOeffnen } from '$lib/plattform';
	import type { VeranstaltungsStore } from '$lib/stores/veranstaltung.svelte';
	import Dialog from '$lib/ui/Dialog.svelte';
	import { ui } from '$lib/ui/ui-zustand.svelte';

	interface Props {
		store: VeranstaltungsStore;
		offen: boolean;
		onschliessen: () => void;
	}

	let { store, offen, onschliessen }: Props = $props();

	let dateiname = $state('');
	let zeilen = $state<NennungsZeile[]>([]);
	let mitKlasse = $state(false);
	let zielKlasse = $state<string>('datei');
	let posten = $state<ImportPosten<Fahrer>[]>([]);
	let datenbank: Fahrer[] = [];
	let lesefehler = $state<string[]>([]);
	let offenerKonflikt = $state<number | null>(null);
	let laeuft = $state(false);

	$effect(() => {
		if (!offen) {
			dateiname = '';
			zeilen = [];
			posten = [];
			lesefehler = [];
			offenerKonflikt = null;
		}
	});

	const artText = { neu: 'Neu in Datenbank', bekannt: 'Bekannt', konflikt: 'Abweichung', 'bereits-gemeldet': 'Bereits gemeldet' } as const;
	const artStil = {
		neu: 'bg-info-soft text-info',
		bekannt: 'bg-ok-soft text-ok',
		konflikt: 'bg-warn-soft text-warn',
		'bereits-gemeldet': 'bg-sunken text-muted'
	} as const;
	const loesungText: Record<KonfliktLoesung, string> = {
		zusammenfuehren: 'Zusammenführen',
		bestehend: 'Datenbank beibehalten',
		'neuer-fahrer': 'Anderer Fahrer (neu anlegen)'
	};

	const zusammenfassung = $derived({
		gesamt: posten.filter((p) => p.uebernehmen).length,
		neu: posten.filter((p) => p.uebernehmen && (p.art === 'neu' || (p.art === 'konflikt' && p.loesung === 'neuer-fahrer'))).length,
		konflikte: posten.filter((p) => p.art === 'konflikt').length
	});

	async function dateiWaehlen() {
		try {
			const datei = await dateiOeffnen('Nennliste importieren', [{ name: 'Nennliste (CSV, Excel)', endungen: ['csv', 'txt', 'xlsx', 'xls', 'ods'] }]);
			if (!datei) return;
			const ergebnis = await leseNennungsDatei(datei.name, datei.bytes);
			datenbank = await (await repo()).fahrerListe();
			dateiname = datei.name;
			zeilen = ergebnis.zeilen;
			lesefehler = ergebnis.fehler;
			mitKlasse = ergebnis.mitKlasse;
			zielKlasse = ergebnis.mitKlasse ? 'datei' : String(store.klassen[0]?.id ?? '');
			planen();
			if (!zeilen.length) ui.melden('In der Datei wurden keine Fahrer gefunden.', 'fehler', lesefehler);
		} catch (e) {
			ui.fehler(e, 'Datei konnte nicht gelesen werden');
		}
	}

	function planen() {
		const ziel = zielKlasse === 'datei' ? null : Number(zielKlasse);
		posten = importPlanen(zeilen, datenbank, store.starter, store.klassen, ziel);
		offenerKonflikt = null;
	}

	function nummernNeu() {
		posten = startnummernVergeben(posten, store.starter);
	}

	function klasseAendern(i: number, wert: string) {
		posten[i].klasseId = wert ? Number(wert) : null;
		if (posten[i].klasseId !== null && posten[i].art !== 'bereits-gemeldet') {
			posten[i].uebernehmen = true;
			posten[i].hinweis = '';
		}
		nummernNeu();
	}

	function kandidatWaehlen(i: number, id: number) {
		const p = posten[i];
		posten[i] = mitBestandVergleichen(p, p.kandidaten.find((k) => k.id === id) ?? null);
	}

	async function importieren() {
		const unklar = posten.filter((p) => p.uebernehmen && p.klasseId === null);
		if (unklar.length) return ui.melden(`${unklar.length} Zeilen ohne Klasse – bitte zuordnen oder abwählen.`, 'warnung');
		laeuft = true;
		try {
			const r = await store.nennungenImportieren($state.snapshot(posten) as ImportPosten<Fahrer>[]);
			ui.melden(
				`${r.genannt} Fahrer gemeldet · ${r.neueFahrer} neu in der Fahrerdatenbank · ${r.aktualisiert} Datensätze zusammengeführt.`,
				r.fehler.length ? 'warnung' : 'erfolg',
				r.fehler
			);
			onschliessen();
		} catch (e) {
			ui.fehler(e, 'Import fehlgeschlagen');
		} finally {
			laeuft = false;
		}
	}

	const anzeige = (wert: unknown, feld: string) => (feld === 'geburtsdatum' ? formatDatum(String(wert ?? '')) : String(wert ?? '')) || '–';
</script>

<Dialog {offen} titel="Nennliste importieren" breite="max-w-6xl" {onschliessen}>
	<div class="flex flex-col gap-4">
		<div class="flex flex-wrap items-end gap-3">
			<button class="btn" onclick={dateiWaehlen}><FileUp size={16} /> {dateiname ? 'Andere Datei wählen' : 'CSV- oder Excel-Datei wählen'}</button>
			{#if dateiname}
				<span class="text-sm text-muted">{dateiname} · {zeilen.length} Fahrer</span>
				<div class="ml-auto">
					<label class="label" for="ziel-klasse">Klasse</label>
					<select id="ziel-klasse" class="input w-56" bind:value={zielKlasse} onchange={planen}>
						{#if mitKlasse}<option value="datei">aus Spalte „Klasse“ der Datei</option>{/if}
						{#each store.klassen as k (k.id)}<option value={String(k.id)}>alle in {k.name}</option>{/each}
					</select>
				</div>
			{/if}
		</div>

		{#if !dateiname}
			<p class="text-sm text-muted">
				Spalten wie in der ZP-Fahrerliste (ID/Lizenz, Klasse, Nachname, Vorname, Rookie, PLZ, Wohnort, Verein, Geburtsdatum), optional „Startnummer“.
				Fahrer, die noch nicht in der Datenbank sind, werden automatisch ergänzt. Abweichende Datensätze werden zum Abgleich angezeigt.
			</p>
		{/if}

		{#if lesefehler.length}
			<details class="rounded-lg bg-warn-soft px-3 py-2 text-sm text-warn">
				<summary>{lesefehler.length} Zeilen konnten nicht gelesen werden</summary>
				<ul class="mt-1 list-disc pl-5 text-xs">{#each lesefehler as f, i (i)}<li>{f}</li>{/each}</ul>
			</details>
		{/if}

		{#if posten.length}
			<div class="overflow-hidden rounded-xl border border-line">
				<table class="w-full text-sm">
					<thead class="bg-sunken text-left text-xs tracking-wide text-muted uppercase">
						<tr>
							<th class="w-8 px-3 py-2"><span class="sr-only">Übernehmen</span></th>
							<th class="px-3 py-2">Nr.</th>
							<th class="px-3 py-2">Fahrer</th>
							<th class="px-3 py-2">Verein</th>
							<th class="px-3 py-2">Klasse</th>
							<th class="px-3 py-2">Abgleich</th>
						</tr>
					</thead>
					<tbody>
						{#each posten as p, i (i)}
							<tr class="border-t border-line align-top {p.uebernehmen ? '' : 'opacity-60'}">
								<td class="px-3 py-2">
									<input
										type="checkbox"
										class="size-4 accent-[var(--color-accent)]"
										bind:checked={p.uebernehmen}
										onchange={nummernNeu}
										disabled={p.art === 'bereits-gemeldet'}
										aria-label="{p.zeile.nachname} übernehmen"
									/>
								</td>
								<td class="px-3 py-2 font-semibold tabular">
									{p.startnummer ?? '–'}
									{#if p.nummerHinweis}<div class="text-[11px] font-normal text-warn">{p.nummerHinweis}</div>{/if}
								</td>
								<td class="px-3 py-2">
									<span class="font-medium">{p.zeile.nachname}, {p.zeile.vorname}</span>
									<div class="font-mono text-xs text-muted">{p.zeile.lizenz || 'ohne Lizenz'}</div>
								</td>
								<td class="px-3 py-2 text-muted">{p.zeile.verein}</td>
								<td class="px-3 py-2">
									<select class="input w-36 py-1" value={p.klasseId === null ? '' : String(p.klasseId)} onchange={(e) => klasseAendern(i, e.currentTarget.value)} aria-label="Klasse">
										<option value="">– wählen –</option>
										{#each store.klassen as k (k.id)}<option value={String(k.id)}>{k.name}</option>{/each}
									</select>
								</td>
								<td class="px-3 py-2">
									<span class="badge {artStil[p.art]}">{artText[p.art]}</span>
									{#if p.art === 'konflikt'}
										<button class="btn btn-sm ml-1" onclick={() => (offenerKonflikt = offenerKonflikt === i ? null : i)}>
											<GitMerge size={13} /> {loesungText[p.loesung]}
										</button>
									{/if}
									{#if p.hinweis && p.art !== 'bereits-gemeldet'}<div class="mt-1 text-xs text-warn">{p.hinweis}</div>{/if}
								</td>
							</tr>
							{#if offenerKonflikt === i && p.bestand}
								{@const ergebnis = zusammengefuehrt(p)}
								<tr class="border-t border-line bg-sunken/40">
									<td colspan="6" class="px-4 py-4">
										<div class="flex flex-wrap items-center gap-2">
											<p class="mr-auto text-sm font-semibold">Abgleich mit der Fahrerdatenbank</p>
											{#if p.kandidaten.length > 1}
												<label class="flex items-center gap-2 text-xs text-muted">
													Vergleichen mit
													<select class="input w-auto py-1" value={String(p.bestand.id)} onchange={(e) => kandidatWaehlen(i, Number(e.currentTarget.value))}>
														{#each p.kandidaten as k (k.id)}<option value={String(k.id)}>{k.nachname}, {k.vorname} ({k.verein})</option>{/each}
													</select>
												</label>
											{/if}
										</div>
										<div class="mt-3 grid grid-cols-[140px_1fr_1fr] overflow-hidden rounded-lg border border-line bg-surface text-sm">
											<div class="bg-sunken px-3 py-2 text-xs font-semibold text-muted uppercase">Feld</div>
											<div class="bg-sunken px-3 py-2 text-xs font-semibold text-muted uppercase">Datenbank</div>
											<div class="bg-sunken px-3 py-2 text-xs font-semibold text-muted uppercase">Nennliste</div>
											{#each ['lizenz', 'nachname', 'vorname', 'verein', 'plz', 'ort', 'geburtsdatum'] as const as feld (feld)}
												{@const abweichend = (p.unterschiede as string[]).includes(feld)}
												<div class="border-t border-line px-3 py-1.5 text-muted">{FELD_NAMEN[feld]}</div>
												{#if abweichend && p.loesung === 'zusammenfuehren'}
													{#each ['bestand', 'import'] as const as seite (seite)}
														<label class="flex cursor-pointer items-center gap-2 border-t border-line px-3 py-1.5 {p.auswahl[feld as keyof typeof p.auswahl] === seite ? 'bg-accent-soft font-semibold' : ''}">
															<input type="radio" name="feld-{i}-{feld}" value={seite} bind:group={p.auswahl[feld as keyof typeof p.auswahl]} class="accent-[var(--color-accent)]" />
															{anzeige(seite === 'bestand' ? p.bestand[feld] : p.zeile[feld], feld)}
														</label>
													{/each}
												{:else}
													<div class="border-t border-line px-3 py-1.5 {abweichend ? 'bg-warn-soft/50' : ''}">{anzeige(p.bestand[feld], feld)}</div>
													<div class="border-t border-line px-3 py-1.5 {abweichend ? 'bg-warn-soft/50' : ''}">{anzeige(p.zeile[feld], feld)}</div>
												{/if}
											{/each}
										</div>
										<div class="mt-3 flex flex-wrap gap-2" role="group" aria-label="Lösung">
											<button class="chip" aria-pressed={p.loesung === 'zusammenfuehren'} onclick={() => (p.loesung = 'zusammenfuehren')}>
												<GitMerge size={14} /> Zusammenführen (gewählte Werte)
											</button>
											<button class="chip" aria-pressed={p.loesung === 'bestehend'} onclick={() => (p.loesung = 'bestehend')}>
												<Users size={14} /> Datenbank beibehalten
											</button>
											<button class="chip" aria-pressed={p.loesung === 'neuer-fahrer'} onclick={() => (p.loesung = 'neuer-fahrer')}>
												<UserPlus size={14} /> Anderer Fahrer – neu anlegen
											</button>
										</div>
										<p class="mt-2 text-xs text-muted">
											{#if p.loesung === 'zusammenfuehren'}
												Der Datensatz wird aktualisiert zu: {ergebnis.nachname}, {ergebnis.vorname} · {ergebnis.verein} · {ergebnis.plz} {ergebnis.ort}. Der bisherige Stand bleibt als Version erhalten; frühere Veranstaltungen behalten ihre Version.
											{:else if p.loesung === 'bestehend'}
												Der Fahrer wird mit den Daten aus der Datenbank gemeldet, die Nennliste wird ignoriert.
											{:else}
												Es wird ein zusätzlicher Fahrer mit den Daten der Nennliste angelegt (gleiche Lizenz, eigener Datensatz).
											{/if}
										</p>
									</td>
								</tr>
							{/if}
						{/each}
					</tbody>
				</table>
			</div>
			{#if zusammenfassung.konflikte}
				<p class="flex items-center gap-2 text-sm text-warn">
					<TriangleAlert size={16} />
					{zusammenfassung.konflikte} Fahrer weichen von der Datenbank ab. Standard ist „Zusammenführen“ mit den Werten der Nennliste – über die Schaltfläche in der Zeile prüfen und anpassen.
				</p>
			{/if}
		{/if}
	</div>
	{#snippet aktionen()}
		{#if posten.length}
			<span class="mr-auto self-center text-sm text-muted">{zusammenfassung.gesamt} Fahrer werden gemeldet, davon {zusammenfassung.neu} neu in der Datenbank</span>
		{/if}
		<button class="btn" onclick={onschliessen}>Abbrechen</button>
		<button class="btn btn-primary" onclick={importieren} disabled={laeuft || zusammenfassung.gesamt === 0}>
			<UserPlus size={16} /> Importieren
		</button>
	{/snippet}
</Dialog>
