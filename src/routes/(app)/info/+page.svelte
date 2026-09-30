<script lang="ts">
	import Seitenkopf from '$lib/components/Seitenkopf.svelte';
	import { istDesktop } from '$lib/plattform';
	import { BUILD_ZEIT, COMMIT, COMMIT_KURZ, VERSION } from '$lib/version';

	const tasten = [
		['Strg + 0 / 1 / 2', 'Erfassung: Training / Lauf 1 / Lauf 2 wählen'],
		['Strg + T', 'Erfassung: nächste freie Zeit aus der Zeitmessung übernehmen'],
		['Strg + S', 'Erfassung: speichern und zum nächsten Start'],
		['↵ (Enter)', 'Erfassung: zum nächsten Feld, im letzten Feld speichern und weiter'],
		['Esc', 'Erfassung: Eingabe verwerfen'],
		['↑ / ↓ und ↵', 'Nennung: Fahrer in der Suche auswählen']
	];

	const ablauf = [
		['Fahrerdatenbank füllen', 'ZP-Fahrerliste (CSV) importieren. Vorhandene Fahrer werden abgeglichen, nicht doppelt angelegt.'],
		['Veranstaltung anlegen', 'Einstellungen, Klassen und Logos werden von der letzten Veranstaltung übernommen.'],
		['Fahrer nennen', 'Fahrer suchen und mit Enter nennen – oder eine Nennliste (CSV/Excel) je Klasse importieren. Gleiche Lizenz oder gleicher Name wie in der Datenbank: abweichende Felder im Seitenvergleich einzeln entscheiden.'],
		['Läufe erfassen', 'Fehler → Fehler → Zeit, jeweils mit Enter; danach springt die Maske zum nächsten Start – Klasse für Klasse: je zwei Fahrer Training und Wertung 1, danach alle Fahrer der Klasse Wertung 2. DNS/DSQ mit Kommentar, Zeiten der Zeitmessung per Klick oder Strg + T.'],
		['Korrigieren', 'Erfasste Läufe in der Erfassung oder über den Stift in der Ergebnisliste ändern – immer mit Begründung, alle Änderungen werden protokolliert.'],
		['Auswerten, drucken, melden', 'Ergebnisse und Mannschaftswertung werden laufend berechnet. Start- und Ergebnislisten, Urkunden und die ZP-Datei unter „Drucken“.']
	];

	const buildZeit = new Date(BUILD_ZEIT).toLocaleString('de-DE', { dateStyle: 'medium', timeStyle: 'short' });
</script>

<Seitenkopf titel="Hilfe & Info" untertitel="Auswertung Light {VERSION} – Auswertprogramm für den Kart-Slalom" />

<div class="grid gap-6 px-7 py-6 lg:grid-cols-[minmax(0,1.2fr)_minmax(0,1fr)]">
	<section class="card flex flex-col gap-3.5 p-6">
		<h2 class="section-title text-[26px]">Ablauf einer Veranstaltung</h2>
		{#each ablauf as [titel, text], i (titel)}
			<div class="flex gap-4">
				<span class="display flex size-12 shrink-0 items-center justify-center rounded-[10px] text-[28px] text-on-ink {i === 3 ? 'bg-accent' : 'bg-ink'}">{i + 1}</span>
				<div class="flex flex-col gap-0.5">
					<span class="text-lg font-bold">{titel}</span>
					<span class="text-[15px] leading-snug text-muted">{text}</span>
				</div>
			</div>
		{/each}
	</section>

	<div class="flex flex-col gap-4">
		<section class="card-ink flex flex-col gap-2.5 p-5">
			<h2 class="section-title text-2xl">Tastenkürzel</h2>
			{#each tasten as [taste, text] (taste)}
				<div class="flex items-center gap-3.5">
					<kbd class="display min-w-36 rounded-md border border-ink-line bg-white/5 px-2.5 py-1 text-center text-lg normal-case">{taste}</kbd>
					<span class="text-[15px] text-ink-muted">{text}</span>
				</div>
			{/each}
		</section>

		<section class="card p-5">
			<h2 class="section-title text-2xl">Wertungsregeln</h2>
			<ul class="mt-2 flex flex-col gap-1.5 text-[15px] leading-snug">
				<li>Laufergebnis = Zeit + Fehler × Strafsekunden (Standard: Pylone 2 s, Tor 10 s).</li>
				<li>Gesamt = Wertungslauf 1 + Wertungslauf 2. Das Training zählt nicht.</li>
				<li>Bei gleicher Gesamtzeit entscheidet der bessere Einzellauf, sonst gleicher Platz.</li>
				<li>Punkte = (Teilnehmer − Platz) × 10 / Teilnehmer + 1.</li>
				<li>Sportabzeichenpunkte: Platz 1 = 6, Platz 2–10 = (12 − Platz) / 2, ab Platz 11 = 0,5.</li>
				<li>Mannschaft: die besten 6 Punktergebnisse je Verein über alle Klassen.</li>
				<li class="text-muted">Fahrer „außer Wertung“ (niW), mit DNS/DSQ in einem Wertungslauf oder ohne beide Wertungsläufe erhalten keinen Platz.</li>
			</ul>
		</section>

		<section class="card p-5">
			<h2 class="section-title text-2xl">Über</h2>
			<dl class="mt-2 grid grid-cols-[auto_1fr] gap-x-4 gap-y-1 text-[15px]">
				<dt class="text-muted">Version</dt>
				<dd class="font-mono text-sm leading-6">{VERSION}</dd>
				<dt class="text-muted">Commit</dt>
				<dd class="font-mono text-sm leading-6 select-all" title={COMMIT}>{COMMIT_KURZ || 'unbekannt'}</dd>
				<dt class="text-muted">Erstellt</dt>
				<dd class="text-sm leading-6">{buildZeit}</dd>
			</dl>
			<p class="mt-3 text-sm text-muted">
				Auswertung Light wurde als Excel-Arbeitsmappe von Johannes Haseitl (derhasi.de) für den Zugspitzpokal entwickelt – mit Unterstützung von Michael Steinhoff und Dieter Schweingruber
				(MC Dießen). Version 2 ist eine Neuentwicklung als Desktop-App.
			</p>
			<p class="mt-2 text-sm text-muted">
				Datenablage: {istDesktop() ? 'SQLite-Datenbank „auswertung-light.db“ im App-Datenverzeichnis des Benutzers.' : 'Browser-Speicher (Entwicklungs- und Demomodus).'}
			</p>
			<p class="mt-2 text-sm">Fragen & Fehler: <span class="font-mono text-xs">github.com/derhasi/auswertung-light/issues</span></p>
		</section>
	</div>
</div>
