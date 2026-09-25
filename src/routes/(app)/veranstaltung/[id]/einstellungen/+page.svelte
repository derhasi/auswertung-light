<script lang="ts">
	import { goto } from '$app/navigation';
	import { untrack } from 'svelte';
	import { ArrowDown, ArrowUp, FileSearch, ImagePlus, Plus, Save, Trash2, X } from '@lucide/svelte';
	import { repo } from '$lib/db';
	import { formatZeit } from '$lib/domain/zahlen';
	import { leseZeitquelle, type GemesseneZeit, type ZeitquelleEinstellung } from '$lib/domain/zeitquelle';
	import { veranstaltungVergessen } from '$lib/stores/veranstaltung.svelte';
	import { bildAlsDataUrl, dateiLesen, dateiOeffnen, istDesktop, pfadWaehlen } from '$lib/plattform';
	import { ui } from '$lib/ui/ui-zustand.svelte';

	let { data } = $props();
	const s = $derived(data.store);

	// Lokale Arbeitskopien der Formulare (Startwerte beim Öffnen der Seite)
	const v0 = untrack(() => data.store.v);
	let stamm = $state({
		name: v0.name,
		datum: v0.datum,
		ort: v0.ort,
		ausrichter: v0.ausrichter,
		zpId: v0.zpId
	});
	let regeln = $state({
		fehler1Name: v0.fehler1Name,
		strafe1: v0.strafe1,
		fehler2Name: v0.fehler2Name,
		strafe2: v0.strafe2,
		mannschaftAnzahl: v0.mannschaftAnzahl,
		urkundenPlaetze: v0.urkundenPlaetze
	});
	let zeitquelle = $state<ZeitquelleEinstellung>({ ...v0.zeitquelle });
	let vorschau = $state<GemesseneZeit[] | null>(null);

	async function speichern(e: SubmitEvent, werte: Record<string, unknown>) {
		e.preventDefault();
		try {
			await s.aktualisieren($state.snapshot(werte));
			ui.melden('Einstellungen gespeichert.');
		} catch (err) {
			ui.fehler(err);
		}
	}

	async function logo(seite: 'logoLinks' | 'logoRechts', entfernen = false) {
		try {
			const bild = entfernen ? '' : await bildAlsDataUrl();
			if (bild === null) return;
			await s.aktualisieren({ [seite]: bild });
		} catch (e) {
			ui.fehler(e);
		}
	}

	async function klasseLoeschen(id: number) {
		try {
			await s.klasseLoeschen(id);
		} catch (e) {
			ui.fehler(e);
		}
	}

	async function zeitquelleWaehlen() {
		const pfad = await pfadWaehlen('Datei der Zeitmessung', [{ name: 'Zeitmessung', endungen: ['csv', 'txt', 'xlsx', 'xls', 'ods'] }]);
		if (pfad) zeitquelle.pfad = pfad;
	}

	async function zeitquelleTesten() {
		try {
			let bytes: Uint8Array;
			if (istDesktop()) {
				bytes = await dateiLesen(zeitquelle.pfad);
			} else {
				const datei = await dateiOeffnen('Testdatei', [{ name: 'Zeitmessung', endungen: ['csv', 'txt', 'xlsx', 'xls', 'ods'] }]);
				if (!datei) return;
				zeitquelle.pfad = datei.name;
				bytes = datei.bytes;
			}
			vorschau = await leseZeitquelle(bytes, $state.snapshot(zeitquelle));
		} catch (e) {
			ui.fehler(e, 'Datei konnte nicht gelesen werden');
		}
	}

	async function veranstaltungLoeschen() {
		const ok = await ui.bestaetigen(
			`„${s.v.name}" mit allen ${s.starter.length} Nennungen und Ergebnissen endgültig löschen?\n\nTipp: Vorher unter „Drucken & Export" eine Sicherung speichern.`,
			{ titel: 'Veranstaltung löschen', ja: 'Endgültig löschen', gefaehrlich: true }
		);
		if (!ok) return;
		await (await repo()).veranstaltungLoeschen(s.id);
		veranstaltungVergessen(s.id);
		ui.melden('Veranstaltung gelöscht.');
		goto('/');
	}
</script>

<div class="grid max-w-5xl gap-6 px-8 py-6">
	<form class="card p-6" onsubmit={(e) => speichern(e, stamm)}>
		<h2 class="mb-4 font-semibold">Veranstaltung</h2>
		<div class="grid gap-4 md:grid-cols-2">
			<div class="md:col-span-2"><label class="label" for="e-name">Bezeichnung</label><input id="e-name" class="input" bind:value={stamm.name} required /></div>
			<div><label class="label" for="e-datum">Datum</label><input id="e-datum" class="input" type="date" bind:value={stamm.datum} required /></div>
			<div><label class="label" for="e-ort">Ort</label><input id="e-ort" class="input" bind:value={stamm.ort} /></div>
			<div><label class="label" for="e-ausrichter">Ausrichter</label><input id="e-ausrichter" class="input" bind:value={stamm.ausrichter} /></div>
			<div><label class="label" for="e-zp">ZP-Veranstaltungs-ID</label><input id="e-zp" class="input font-mono" bind:value={stamm.zpId} /></div>
		</div>
		<p class="mt-2 text-xs text-muted">Das Datum bestimmt das Rookie-Jahr: Als Rookie gilt, wessen Rookie-Jahr dem Veranstaltungsjahr entspricht.</p>
		<div class="mt-4 flex justify-end"><button class="btn btn-primary"><Save size={16} /> Speichern</button></div>
	</form>

	<form class="card p-6" onsubmit={(e) => speichern(e, regeln)}>
		<h2 class="mb-4 font-semibold">Wertung</h2>
		<div class="grid gap-4 md:grid-cols-4">
			<div><label class="label" for="r-f1">Fehlerart 1</label><input id="r-f1" class="input" bind:value={regeln.fehler1Name} required /></div>
			<div><label class="label" for="r-s1">Strafsekunden</label><input id="r-s1" class="input" type="number" min="0" step="0.01" bind:value={regeln.strafe1} required /></div>
			<div><label class="label" for="r-f2">Fehlerart 2</label><input id="r-f2" class="input" bind:value={regeln.fehler2Name} required /></div>
			<div><label class="label" for="r-s2">Strafsekunden</label><input id="r-s2" class="input" type="number" min="0" step="0.01" bind:value={regeln.strafe2} required /></div>
			<div class="md:col-span-2">
				<label class="label" for="r-m">Wertende Ergebnisse je Verein (Mannschaft)</label>
				<input id="r-m" class="input" type="number" min="1" max="50" bind:value={regeln.mannschaftAnzahl} required />
			</div>
			<div class="md:col-span-2">
				<label class="label" for="r-u">Urkunden für die Plätze 1 bis</label>
				<input id="r-u" class="input" type="number" min="1" max="50" bind:value={regeln.urkundenPlaetze} required />
			</div>
		</div>
		<div class="mt-4 flex justify-end"><button class="btn btn-primary"><Save size={16} /> Speichern</button></div>
	</form>

	<section class="card p-6">
		<div class="mb-4 flex items-center justify-between">
			<h2 class="font-semibold">Klassen</h2>
			<button class="btn btn-sm" onclick={() => s.klasseAnlegen().catch((e) => ui.fehler(e))}><Plus size={14} /> Klasse hinzufügen</button>
		</div>
		<ul class="divide-y divide-line rounded-lg border border-line">
			{#each s.klassen as k, i (k.id)}
				{@const anzahl = s.starter.filter((st) => st.klasseId === k.id).length}
				<li class="flex flex-wrap items-center gap-3 px-3 py-2">
					<div class="flex flex-col">
						<button class="text-muted hover:text-fg disabled:opacity-30" disabled={i === 0} onclick={() => s.klasseVerschieben(k.id, -1)} aria-label="Nach oben"><ArrowUp size={14} /></button>
						<button class="text-muted hover:text-fg disabled:opacity-30" disabled={i === s.klassen.length - 1} onclick={() => s.klasseVerschieben(k.id, 1)} aria-label="Nach unten"><ArrowDown size={14} /></button>
					</div>
					<input class="input w-48" value={k.name} aria-label="Name" onchange={(e) => s.klasseAktualisieren(k.id, { name: e.currentTarget.value })} />
					<input class="input w-24 font-mono" value={k.kuerzel} aria-label="Kürzel" title="Kürzel wie in der Fahrerliste (z. B. K1)" onchange={(e) => s.klasseAktualisieren(k.id, { kuerzel: e.currentTarget.value })} />
					<label class="flex items-center gap-2 text-sm">
						<input type="checkbox" class="size-4 accent-[var(--color-accent)]" checked={k.inMannschaft} onchange={(e) => s.klasseAktualisieren(k.id, { inMannschaft: e.currentTarget.checked })} />
						zählt zur Mannschaftswertung
					</label>
					<span class="ml-auto text-xs text-muted">{anzahl} Starter</span>
					<button class="btn btn-ghost btn-icon text-muted hover:text-danger" disabled={anzahl > 0} title={anzahl > 0 ? 'Klasse enthält noch Starter' : 'Klasse löschen'} onclick={() => klasseLoeschen(k.id)} aria-label="Klasse löschen"><Trash2 size={16} /></button>
				</li>
			{/each}
		</ul>
		<p class="mt-2 text-xs text-muted">Beim Nennen wird die Klasse aus der Fahrerdatenbank über das Kürzel zugeordnet (z. B. „K1" → Klasse 1).</p>
	</section>

	<section class="card p-6">
		<h2 class="mb-4 font-semibold">Logos für Ausdrucke</h2>
		<div class="grid gap-6 md:grid-cols-2">
			{#each [['logoLinks', 'Logo links'], ['logoRechts', 'Logo rechts']] as const as [feld, titel] (feld)}
				<div>
					<p class="label">{titel}</p>
					<div class="flex h-28 items-center justify-center rounded-lg border border-dashed border-line bg-sunken/50 p-2">
						{#if s.v[feld]}<img src={s.v[feld]} alt={titel} class="max-h-full max-w-full object-contain" />{:else}<span class="text-sm text-muted">kein Logo</span>{/if}
					</div>
					<div class="mt-2 flex gap-2">
						<button class="btn btn-sm" onclick={() => logo(feld)}><ImagePlus size={14} /> Bild wählen</button>
						{#if s.v[feld]}<button class="btn btn-sm btn-ghost" onclick={() => logo(feld, true)}><X size={14} /> Entfernen</button>{/if}
					</div>
				</div>
			{/each}
		</div>
	</section>

	<form class="card p-6" onsubmit={(e) => speichern(e, { zeitquelle })}>
		<h2 class="font-semibold">Zeitmessung</h2>
		<p class="mt-1 mb-4 text-sm text-muted">
			Die Zeiten können aus einer CSV- oder Excel-Datei der Zeitmessanlage übernommen werden. Die Desktop-App beobachtet die Datei und zeigt neue Zeiten in der Erfassung sofort an.
		</p>
		<div class="grid gap-4 md:grid-cols-6">
			<div class="md:col-span-6">
				<label class="label" for="z-pfad">Datei</label>
				<div class="flex gap-2">
					<input id="z-pfad" class="input font-mono text-xs" bind:value={zeitquelle.pfad} placeholder={istDesktop() ? 'C:\\Zeitmessung\\zeiten.csv' : 'Dateiauswahl nur in der Desktop-App'} readonly={!istDesktop()} />
					{#if istDesktop()}<button type="button" class="btn" onclick={zeitquelleWaehlen}>Durchsuchen …</button>{/if}
				</div>
			</div>
			<div class="md:col-span-2"><label class="label" for="z-blatt">Tabellenblatt (Excel)</label><input id="z-blatt" class="input" bind:value={zeitquelle.blatt} placeholder="erstes Blatt" /></div>
			<div><label class="label" for="z-id">Spalte Kennung</label><input id="z-id" class="input" type="number" min="0" bind:value={zeitquelle.idSpalte} /></div>
			<div><label class="label" for="z-zeit">Spalte Zeit</label><input id="z-zeit" class="input" type="number" min="1" bind:value={zeitquelle.zeitSpalte} /></div>
			<div><label class="label" for="z-kopf">Kopfzeilen</label><input id="z-kopf" class="input" type="number" min="0" bind:value={zeitquelle.kopfzeilen} /></div>
			<div>
				<label class="label" for="z-format">Format</label>
				<select id="z-format" class="input" bind:value={zeitquelle.format}>
					<option value="dezimal">Sekunden (32,45)</option>
					<option value="zeit">Zeit (0:32,45)</option>
				</select>
			</div>
		</div>
		<p class="mt-2 text-xs text-muted">Spalten zählen ab 1 (A = 1, B = 2 …). Kennung 0 = Zeilennummer der Datei.</p>
		{#if vorschau}
			<div class="mt-4 rounded-lg border border-line p-3 text-sm">
				<p class="font-medium">{vorschau.length} Zeiten gefunden{vorschau.length ? ' – die letzten:' : '.'}</p>
				<ul class="mt-1 flex flex-wrap gap-2">
					{#each vorschau.slice(-8) as z (z.zeile)}<li class="rounded bg-sunken px-2 py-0.5 font-mono text-xs">#{z.id}: {formatZeit(z.zeit)}</li>{/each}
				</ul>
			</div>
		{/if}
		<div class="mt-4 flex justify-end gap-2">
			<button type="button" class="btn" onclick={zeitquelleTesten} disabled={istDesktop() && !zeitquelle.pfad}><FileSearch size={16} /> Testen</button>
			<button class="btn btn-primary"><Save size={16} /> Speichern</button>
		</div>
	</form>

	<section class="card border-danger/40 p-6">
		<h2 class="font-semibold text-danger">Veranstaltung löschen</h2>
		<p class="mt-1 text-sm text-muted">Löscht die Veranstaltung mit allen Nennungen und Ergebnissen. Die Fahrerdatenbank bleibt erhalten.</p>
		<button class="btn btn-danger mt-4" onclick={veranstaltungLoeschen}><Trash2 size={16} /> Veranstaltung löschen</button>
	</section>
</div>
