<script lang="ts">
	import { Award, ClipboardList, FileJson, FileSpreadsheet, Trophy, Upload, Users } from '@lucide/svelte';
	import { repo } from '$lib/db';
	import { kodiereWindows1252, stringifyCsv } from '$lib/domain/csv';
	import { LAUF_KURZ, WERTUNGSLAEUFE } from '$lib/domain/typen';
	import { formatPunkte, formatZeit } from '$lib/domain/zahlen';
	import { zpExportCsv } from '$lib/domain/zp-export';
	import { CSV_FILTER, dateiSpeichern, JSON_FILTER } from '$lib/plattform';
	import { ui } from '$lib/ui/meldungen.svelte';

	let { data } = $props();
	const s = $derived(data.store);

	const unvollstaendig = $derived(s.klassenWertungen.flatMap((k) => k.zeilen.filter((z) => z.status === 'unvollstaendig')));
	const ohneLizenz = $derived(s.starter.filter((st) => !st.lizenz));
	const dateiBasis = $derived(`${s.v.datum}_${s.v.name}`.replace(/[^\p{L}\p{N}_-]+/gu, '_'));

	const druckstuecke = [
		{ dok: 'start', titel: 'Startliste', text: 'Nach Startnummer, mit Feldern für Notizen', icon: ClipboardList },
		{ dok: 'ergebnis', titel: 'Ergebnisliste', text: 'Je Klasse mit Laufzeiten, Punkten und Adressen', icon: Trophy },
		{ dok: 'mannschaft', titel: 'Mannschaftswertung', text: 'Vereinswertung mit wertenden Fahrern', icon: Users },
		{ dok: 'urkunden', titel: 'Urkunden', text: 'Für die Plätze 1 bis {n} jeder Klasse', icon: Award }
	];

	async function zpSpeichern() {
		if (!s.v.zpId.trim()) {
			const ok = await ui.bestaetigen('Es ist keine ZP-Veranstaltungs-ID eingetragen. Die erste Spalte bleibt dann leer. Trotzdem speichern?', { ja: 'Trotzdem speichern' });
			if (!ok) return;
		}
		try {
			const csv = zpExportCsv(s.v.zpId.trim(), s.klassenWertungen);
			const ziel = await dateiSpeichern('zp_output.csv', kodiereWindows1252(csv), [{ name: 'Zugspitzpokal Output', endungen: ['csv'] }]);
			if (ziel) ui.melden('ZP-Output gespeichert.');
		} catch (e) {
			ui.fehler(e, 'Speichern fehlgeschlagen');
		}
	}

	async function ergebnisCsv() {
		const kopf = ['Klasse', 'Platz', 'Startnummer', 'Lizenz', 'Nachname', 'Vorname', 'Verein', 'PLZ', 'Ort', 'Rookie'];
		for (const nr of [0, ...WERTUNGSLAEUFE] as const) kopf.push(`${LAUF_KURZ[nr]} ${s.v.fehler1Name}`, `${LAUF_KURZ[nr]} ${s.v.fehler2Name}`, `${LAUF_KURZ[nr]} Zeit`, `${LAUF_KURZ[nr]} Ergebnis`);
		kopf.push('Gesamt', 'Punkte', 'Sportabzeichenpunkte');
		const zeilen = s.klassenWertungen.flatMap(({ klasse, zeilen }) =>
			zeilen.map((z) => {
				const st = z.starter;
				const werte: (string | number)[] = [
					klasse.name,
					st.ausserWertung ? 'niW' : (z.platz ?? ''),
					st.startnummer, st.lizenz, st.nachname, st.vorname, st.verein, st.plz, st.ort, z.rookie ? 'ja' : ''
				];
				for (const nr of [0, 1, 2] as const) {
					const l = st.laeufe[nr];
					werte.push(l?.fehler1 ?? '', l?.fehler2 ?? '', formatZeit(l?.zeit), formatZeit(z.ergebnisse[nr]));
				}
				werte.push(formatZeit(z.gesamt), z.status === 'gewertet' ? formatPunkte(z.punkte) : '', z.status === 'gewertet' ? formatPunkte(z.sportabzeichen) : '');
				return werte;
			})
		);
		// Excel (deutsch) erwartet Semikolon und erkennt UTF-8 am BOM.
		const ziel = await dateiSpeichern(`${dateiBasis}_Ergebnisse.csv`, '﻿' + stringifyCsv([kopf, ...zeilen], ';'), CSV_FILTER).catch((e) => ui.fehler(e));
		if (ziel) ui.melden('Ergebnisse exportiert.');
	}

	async function sichern() {
		try {
			const daten = await (await repo()).veranstaltungExportieren(s.id);
			const ziel = await dateiSpeichern(`${dateiBasis}.json`, JSON.stringify(daten, null, 2), JSON_FILTER);
			if (ziel) ui.melden('Veranstaltung gesichert. Sie kann auf der Startseite wieder importiert werden.');
		} catch (e) {
			ui.fehler(e);
		}
	}
</script>

<div class="grid gap-8 px-8 py-6">
	<section>
		<h2 class="mb-3 text-sm font-semibold tracking-wide text-muted uppercase">Drucken</h2>
		<div class="grid gap-4 [grid-template-columns:repeat(auto-fill,minmax(230px,1fr))]">
			{#each druckstuecke as d (d.dok)}
				{@const Icon = d.icon}
				<a class="card flex flex-col gap-2 p-5 hover:shadow-md" href="/druck/{s.id}?dok={d.dok}">
					<Icon class="text-accent" />
					<p class="font-semibold">{d.titel}</p>
					<p class="text-sm text-muted">{d.text.replace('{n}', String(s.v.urkundenPlaetze))}</p>
				</a>
			{/each}
		</div>
		<p class="mt-2 text-xs text-muted">Im Druckdialog kann statt eines Druckers auch „Als PDF speichern" gewählt werden.</p>
	</section>

	<section class="card p-6">
		<div class="flex flex-wrap items-start gap-6">
			<div class="min-w-64 flex-1">
				<h2 class="flex items-center gap-2 font-semibold"><Upload size={18} class="text-accent" /> ZP-Export für zugspitzpokal.de</h2>
				<p class="mt-1 text-sm text-muted">
					Erzeugt die Datei <code>zp_output.csv</code> für den Ergebnis- und Statistikdienst (gleiches Format wie das bisherige Blatt „zp_output").
				</p>
				{#if unvollstaendig.length}
					<p class="mt-3 rounded-lg bg-warn-soft px-3 py-2 text-sm text-warn">
						{unvollstaendig.length} Fahrer ohne vollständige Wertungsläufe: {unvollstaendig.map((z) => z.starter.startnummer).join(', ')}
					</p>
				{/if}
				{#if ohneLizenz.length}
					<p class="mt-2 rounded-lg bg-info-soft px-3 py-2 text-sm text-info">{ohneLizenz.length} Fahrer ohne Lizenznummer.</p>
				{/if}
			</div>
			<div class="flex w-72 flex-col gap-3">
				<div>
					<label class="label" for="zp-id">ZP-Veranstaltungs-ID</label>
					<input id="zp-id" class="input font-mono" value={s.v.zpId} onchange={(e) => s.aktualisieren({ zpId: e.currentTarget.value.trim() })} />
				</div>
				<button class="btn btn-primary" onclick={zpSpeichern}><FileSpreadsheet size={16} /> ZP-Datei speichern</button>
			</div>
		</div>
	</section>

	<section>
		<h2 class="mb-3 text-sm font-semibold tracking-wide text-muted uppercase">Weitere Exporte</h2>
		<div class="flex flex-wrap gap-3">
			<button class="btn" onclick={ergebnisCsv}><FileSpreadsheet size={16} /> Ergebnisse als CSV (Excel)</button>
			<button class="btn" onclick={sichern}><FileJson size={16} /> Veranstaltung sichern (JSON)</button>
		</div>
	</section>
</div>
