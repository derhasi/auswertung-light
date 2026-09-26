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

	const buildZeit = new Date(BUILD_ZEIT).toLocaleString('de-DE', { dateStyle: 'medium', timeStyle: 'short' });
</script>

<Seitenkopf titel="Hilfe & Info" untertitel="Auswertung Light {VERSION} – Auswertprogramm für den Kart-Slalom" />

<div class="grid max-w-5xl gap-6 px-8 pb-10 lg:grid-cols-2">
	<section class="card p-6">
		<h2 class="font-semibold">Ablauf einer Veranstaltung</h2>
		<ol class="mt-3 list-decimal space-y-2 pl-5 text-sm">
			<li><strong>Fahrerdatenbank</strong>: ZP-Fahrerliste (CSV) importieren. Vorhandene Fahrer werden abgeglichen, nicht doppelt angelegt.</li>
			<li><strong>Veranstaltung anlegen</strong>: Einstellungen, Klassen und Logos werden von der letzten Veranstaltung übernommen.</li>
			<li><strong>Nennung</strong>: Fahrer suchen und mit Enter nennen – oder eine Nennliste (CSV/Excel) je Klasse importieren. Gleiche Lizenz oder gleicher Name wie in der Datenbank: abweichende Felder im Seitenvergleich einzeln entscheiden.</li>
			<li><strong>Erfassung</strong>: Fehler → Fehler → Zeit, jeweils mit Enter; danach springt die Maske zum nächsten Start – Klasse für Klasse: je zwei Fahrer Training und Wertung 1, danach alle Fahrer der Klasse Wertung 2. Nach dem letzten Lauf einer Klasse: Ergebnis anzeigen oder zur nächsten Klasse. DNS/DSQ mit Kommentar. Zeiten der Zeitmessung per Klick oder Strg + T.</li>
			<li><strong>Korrekturen</strong>: erfasste Läufe in der Erfassung oder über den Stift in der Ergebnisliste ändern – immer mit Begründung, alle Änderungen werden protokolliert.</li>
			<li><strong>Ergebnisse & Mannschaft</strong>: werden laufend berechnet.</li>
			<li><strong>Drucken & Export</strong>: Start- und Ergebnislisten, Mannschaftswertung, Urkunden, ZP-Datei.</li>
		</ol>
	</section>

	<section class="card p-6">
		<h2 class="font-semibold">Tastenkürzel</h2>
		<dl class="mt-3 grid grid-cols-[auto_1fr] gap-x-4 gap-y-2 text-sm">
			{#each tasten as [taste, text] (taste)}
				<dt><kbd class="rounded border border-line bg-sunken px-1.5 py-0.5 font-mono text-xs">{taste}</kbd></dt>
				<dd class="text-muted">{text}</dd>
			{/each}
		</dl>
	</section>

	<section class="card p-6">
		<h2 class="font-semibold">Wertungsregeln</h2>
		<ul class="mt-3 list-disc space-y-1.5 pl-5 text-sm text-muted">
			<li>Laufergebnis = Zeit + Fehler × Strafsekunden (Standard: Pylone 2 s, Tor 10 s).</li>
			<li>Gesamt = Wertungslauf 1 + Wertungslauf 2. Das Training zählt nicht.</li>
			<li>Bei gleicher Gesamtzeit entscheidet der bessere Einzellauf, sonst gleicher Platz.</li>
			<li>Punkte = (Teilnehmer − Platz) × 10 / Teilnehmer + 1.</li>
			<li>Sportabzeichenpunkte: Platz 1 = 6, Platz 2–10 = (12 − Platz) / 2, ab Platz 11 = 0,5.</li>
			<li>Mannschaft: die besten 6 Punktergebnisse je Verein über alle Klassen.</li>
			<li>Fahrer „außer Wertung" (niW), mit DNS/DSQ in einem Wertungslauf oder ohne beide Wertungsläufe erhalten keinen Platz.</li>
		</ul>
	</section>

	<section class="card p-6">
		<h2 class="font-semibold">Über</h2>
		<dl class="mt-3 grid grid-cols-[auto_1fr] gap-x-4 gap-y-1 text-sm">
			<dt class="text-muted">Version</dt>
			<dd class="font-mono text-xs leading-5">{VERSION}</dd>
			<dt class="text-muted">Commit</dt>
			<dd class="font-mono text-xs leading-5 select-all" title={COMMIT}>{COMMIT_KURZ || 'unbekannt'}</dd>
			<dt class="text-muted">Erstellt</dt>
			<dd class="text-xs leading-5">{buildZeit}</dd>
		</dl>
		<p class="mt-3 text-sm text-muted">
			Auswertung Light wurde als Excel-Arbeitsmappe von Johannes Haseitl (derhasi.de) für den Zugspitzpokal entwickelt – mit Unterstützung von Michael Steinhoff und Dieter Schweingruber
			(MC Dießen). Version 2 ist eine Neuentwicklung als Desktop-App.
		</p>
		<p class="mt-3 text-sm text-muted">
			Datenablage: {istDesktop() ? 'SQLite-Datenbank „auswertung-light.db" im App-Datenverzeichnis des Benutzers.' : 'Browser-Speicher (Entwicklungs- und Demomodus).'}
		</p>
		<p class="mt-3 text-sm">
			Fragen & Fehler: <span class="font-mono text-xs">github.com/derhasi/auswertung-light/issues</span>
		</p>
	</section>
</div>
