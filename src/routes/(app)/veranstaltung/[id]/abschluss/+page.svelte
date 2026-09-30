<script lang="ts">
	import { Award, ClipboardList, FileJson, FileSpreadsheet, Printer, Trophy, Users } from '@lucide/svelte';
	import Seitenkopf from '$lib/components/Seitenkopf.svelte';
	import { repo } from '$lib/db';
	import { kodiereWindows1252, stringifyCsv } from '$lib/domain/csv';
	import { LAUF_KURZ, WERTUNGSLAEUFE } from '$lib/domain/typen';
	import { formatPunkte, formatZeit } from '$lib/domain/zahlen';
	import { zpExportCsv } from '$lib/domain/zp-export';
	import { STATUS_KURZ } from '$lib/domain/wertung';
	import { CSV_FILTER, dateiSpeichern, JSON_FILTER } from '$lib/plattform';
	import { ui } from '$lib/ui/ui-zustand.svelte';

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
		for (const nr of [0, ...WERTUNGSLAEUFE] as const)
			kopf.push(`${LAUF_KURZ[nr]} ${s.v.fehler1Name}`, `${LAUF_KURZ[nr]} ${s.v.fehler2Name}`, `${LAUF_KURZ[nr]} Zeit`, `${LAUF_KURZ[nr]} Ergebnis`, `${LAUF_KURZ[nr]} Kommentar`);
		kopf.push('Gesamt', 'Punkte', 'Sportabzeichenpunkte');
		const zeilen = s.klassenWertungen.flatMap(({ klasse, zeilen }) =>
			zeilen.map((z) => {
				const st = z.starter;
				const werte: (string | number)[] = [
					klasse.name,
					z.platz ?? STATUS_KURZ[z.status],
					st.startnummer, st.lizenz, st.nachname, st.vorname, st.verein, st.plz, st.ort, z.rookie ? 'ja' : ''
				];
				for (const nr of [0, 1, 2] as const) {
					const l = st.laeufe[nr];
					const status = (l?.status ?? 'ok') !== 'ok' ? (l?.status ?? '').toUpperCase() : '';
					werte.push(status ? '' : (l?.fehler1 ?? ''), status ? '' : (l?.fehler2 ?? ''), status || formatZeit(l?.zeit), status || formatZeit(z.ergebnisse[nr]), l?.kommentar ?? '');
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

<Seitenkopf titel="Drucken & Export" untertitel="Aushänge, Urkunden und die Meldung an den Zugspitzpokal" />

<div class="grid gap-6 px-7 py-6 lg:grid-cols-[minmax(0,1fr)_380px]">
	<section class="flex flex-col gap-3">
		<div class="grid gap-4 md:grid-cols-2">
			{#each druckstuecke as d (d.dok)}
				{@const Icon = d.icon}
				<div class="card flex flex-col gap-3 p-5">
					<div class="flex items-center gap-3.5">
						<span class="flex size-13 shrink-0 items-center justify-center rounded-[10px] bg-ink text-on-ink"><Icon size={24} /></span>
						<h2 class="display text-[30px] leading-none">{d.titel}</h2>
					</div>
					<p class="text-base text-muted">{d.text.replace('{n}', String(s.v.urkundenPlaetze))}</p>
					<a class="btn btn-primary btn-lg mt-auto self-start" href="/druck/{s.id}?dok={d.dok}"><Printer size={20} /> Drucken / PDF</a>
				</div>
			{/each}
		</div>
		<p class="text-sm text-muted">Im Druckdialog kann statt eines Druckers auch „Als PDF speichern“ gewählt werden.</p>
	</section>

	<aside class="flex flex-col gap-4">
		<section class="card overflow-hidden">
			<h2 class="section-title bg-ink px-5 py-3.5 text-2xl text-on-ink">ZP-Export</h2>
			<div class="flex flex-col gap-3 p-5">
				<div>
					<label class="label" for="zp-id">ZP-Veranstaltungs-ID</label>
					<input id="zp-id" class="input font-mono" value={s.v.zpId} onchange={(e) => s.aktualisieren({ zpId: e.currentTarget.value.trim() })} />
				</div>
				<p class="text-sm text-muted">Erzeugt <code>zp_output.csv</code> für zugspitzpokal.de – gleiches Format wie das bisherige Blatt „zp_output“.</p>
				{#if unvollstaendig.length}
					<p class="rounded-lg bg-warn-soft px-3 py-2 text-sm text-warn">
						{unvollstaendig.length} Fahrer ohne vollständige Wertungsläufe: {unvollstaendig.map((z) => z.starter.startnummer).join(', ')}
					</p>
				{/if}
				{#if ohneLizenz.length}
					<p class="rounded-lg bg-info-soft px-3 py-2 text-sm text-info">{ohneLizenz.length} Fahrer ohne Lizenznummer.</p>
				{/if}
				<button class="btn btn-primary btn-lg" onclick={zpSpeichern}>zp_output.csv speichern</button>
			</div>
		</section>

		<section class="card flex flex-col overflow-hidden">
			<h2 class="eyebrow px-5 pt-4 pb-2 text-muted">Weitere Exporte</h2>
			<button class="flex items-start gap-3 border-t border-line px-5 py-3 text-left hover:bg-sunken" onclick={ergebnisCsv}>
				<FileSpreadsheet size={20} class="mt-0.5 text-accent" />
				<span class="flex flex-col"><span class="font-bold">Ergebnisse als CSV</span><span class="text-sm text-muted">Für Excel, mit Status und Kommentar</span></span>
			</button>
			<button class="flex items-start gap-3 border-t border-line px-5 py-3 text-left hover:bg-sunken" onclick={sichern}>
				<FileJson size={20} class="mt-0.5 text-accent" />
				<span class="flex flex-col"><span class="font-bold">Veranstaltung sichern</span><span class="text-sm text-muted">JSON-Datei, auf der Startseite wieder einlesbar</span></span>
			</button>
		</section>
	</aside>
</div>
