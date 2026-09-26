<script lang="ts">
	import { goto } from '$app/navigation';
	import { CalendarPlus, FileUp, MapPin, Trophy, Users, Flag } from '@lucide/svelte';
	import { repo, type VeranstaltungsUebersicht } from '$lib/db';
	import { formatDatum } from '$lib/domain/fahrer-import';
	import Seitenkopf from '$lib/components/Seitenkopf.svelte';
	import Dialog from '$lib/ui/Dialog.svelte';
	import { ui } from '$lib/ui/ui-zustand.svelte';
	import { dateiOeffnen, JSON_FILTER } from '$lib/plattform';
	import { dekodiereText } from '$lib/domain/csv';

	let liste = $state<VeranstaltungsUebersicht[] | null>(null);
	let fahrerAnzahl = $state(0);
	let dialogOffen = $state(false);
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
			dialogOffen = false;
			goto(`/veranstaltung/${id}/nennung`);
		} catch (e) {
			ui.fehler(e);
		}
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
		<button class="btn" onclick={importieren}><FileUp size={16} /> Importieren</button>
		<button class="btn btn-primary" onclick={() => (dialogOffen = true)}><CalendarPlus size={16} /> Neue Veranstaltung</button>
	{/snippet}
</Seitenkopf>

<div class="px-8 pb-10">
	{#if liste === null}
		<p class="text-sm text-muted">Lade …</p>
	{:else if liste.length === 0}
		<div class="card mx-auto max-w-2xl p-8">
			<h2 class="text-lg font-semibold">Willkommen bei Auswertung Light</h2>
			<p class="mt-1 text-sm text-muted">In drei Schritten zur ersten Auswertung:</p>
			<ol class="mt-6 space-y-4">
				<li class="flex gap-4">
					<span class="flex size-8 shrink-0 items-center justify-center rounded-full bg-accent-soft font-bold text-accent-strong">1</span>
					<div>
						<p class="font-medium">Fahrerdatenbank füllen</p>
						<p class="text-sm text-muted">
							Die Fahrerliste des Zugspitzpokals (CSV) importieren oder Fahrer von Hand anlegen.
							{#if fahrerAnzahl}<span class="badge bg-ok-soft text-ok">{fahrerAnzahl} Fahrer vorhanden</span>{/if}
						</p>
						<a class="btn btn-sm mt-2" href="/fahrer"><Users size={14} /> Zur Fahrerdatenbank</a>
					</div>
				</li>
				<li class="flex gap-4">
					<span class="flex size-8 shrink-0 items-center justify-center rounded-full bg-accent-soft font-bold text-accent-strong">2</span>
					<div>
						<p class="font-medium">Veranstaltung anlegen und Fahrer nennen</p>
						<p class="text-sm text-muted">Klassen, Strafsekunden und Logos lassen sich je Veranstaltung einstellen.</p>
						<button class="btn btn-sm btn-primary mt-2" onclick={() => (dialogOffen = true)}><CalendarPlus size={14} /> Veranstaltung anlegen</button>
					</div>
				</li>
				<li class="flex gap-4">
					<span class="flex size-8 shrink-0 items-center justify-center rounded-full bg-accent-soft font-bold text-accent-strong">3</span>
					<div>
						<p class="font-medium">Läufe erfassen, Ergebnisse drucken, ZP-Export speichern</p>
						<p class="text-sm text-muted">Ergebnisliste und Mannschaftswertung werden live berechnet.</p>
					</div>
				</li>
			</ol>
		</div>
	{:else}
		<div class="grid gap-4 [grid-template-columns:repeat(auto-fill,minmax(280px,1fr))]">
			{#each liste as v (v.id)}
				<a href="/veranstaltung/{v.id}" class="card group flex flex-col p-5 transition-shadow hover:shadow-lg">
					<div class="flex items-center gap-2 text-xs font-semibold text-accent-strong">
						<Flag size={14} />
						{formatDatum(v.datum)}
					</div>
					<h2 class="mt-2 text-lg leading-snug font-semibold group-hover:text-accent-strong">{v.name}</h2>
					{#if v.ort}<p class="mt-1 flex items-center gap-1 text-sm text-muted"><MapPin size={14} />{v.ort}</p>{/if}
					<div class="mt-4 flex gap-4 border-t border-line pt-3 text-sm text-muted">
						<span class="flex items-center gap-1.5"><Users size={14} /> {v.starter} Starter</span>
						<span class="flex items-center gap-1.5"><Trophy size={14} /> {v.klassen} Klassen</span>
					</div>
				</a>
			{/each}
		</div>
	{/if}
</div>

<Dialog offen={dialogOffen} titel="Neue Veranstaltung" onschliessen={() => (dialogOffen = false)}>
	<form id="neu" class="grid gap-4" onsubmit={anlegen}>
		<div>
			<label class="label" for="neu-name">Bezeichnung</label>
			<!-- svelte-ignore a11y_autofocus -->
			<input id="neu-name" class="input" bind:value={neu.name} placeholder="z. B. 3. Lauf Zugspitzpokal" required autofocus />
		</div>
		<div class="grid grid-cols-2 gap-4">
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
				{#each liste ?? [] as v (v.id)}<option value={String(v.id)}>von „{v.name}" ({formatDatum(v.datum)})</option>{/each}
				<option value="keine">keine – Standardeinstellungen (Klasse 1–6, 2 s / 10 s)</option>
			</select>
			<p class="mt-1 text-xs text-muted">Übernommen werden Klassen, Strafsekunden, Logos, Ausrichter und Zeitmessung – keine Fahrer.</p>
		</div>
	</form>
	{#snippet aktionen()}
		<button class="btn" onclick={() => (dialogOffen = false)}>Abbrechen</button>
		<button class="btn btn-primary" type="submit" form="neu">Anlegen</button>
	{/snippet}
</Dialog>
