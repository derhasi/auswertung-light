<script lang="ts">
	import Seitenkopf from '$lib/components/Seitenkopf.svelte';
	import { istDesktop } from '$lib/plattform';
	import { VERSION } from '$lib/version';

	const tasten = [
		['Strg + 0 / 1 / 2', 'Erfassung: Training / Lauf 1 / Lauf 2 wählen'],
		['Strg + T', 'Erfassung: nächste freie Zeit aus der Zeitmessung übernehmen'],
		['Strg + S', 'Erfassung: speichern'],
		['↵ (Enter)', 'Erfassung: zum nächsten Feld, im Zeitfeld speichern'],
		['Esc', 'Erfassung: Eingabe verwerfen'],
		['↑ / ↓ und ↵', 'Nennung: Fahrer in der Suche auswählen']
	];
</script>

<Seitenkopf titel="Hilfe & Info" untertitel="Auswertung Light {VERSION} – Auswertprogramm für den Kart-Slalom" />

<div class="grid max-w-5xl gap-6 px-8 pb-10 lg:grid-cols-2">
	<section class="card p-6">
		<h2 class="font-semibold">Ablauf einer Veranstaltung</h2>
		<ol class="mt-3 list-decimal space-y-2 pl-5 text-sm">
			<li><strong>Fahrerdatenbank</strong>: ZP-Fahrerliste (CSV) importieren. Vorhandene Fahrer werden abgeglichen, nicht doppelt angelegt.</li>
			<li><strong>Veranstaltung anlegen</strong>: Einstellungen, Klassen und Logos werden von der letzten Veranstaltung übernommen.</li>
			<li><strong>Nennung</strong>: Fahrer suchen, Klasse und Startnummer prüfen, mit Enter nennen. Fahrer ohne Lizenz können direkt erfasst werden.</li>
			<li><strong>Erfassung</strong>: Startnummer → Fehler → Zeit, jeweils mit Enter. Zeiten der Zeitmessung lassen sich per Klick oder Strg + T übernehmen.</li>
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
			<li>Fahrer „außer Wertung" (niW) und Fahrer ohne beide Wertungsläufe erhalten keinen Platz.</li>
		</ul>
	</section>

	<section class="card p-6">
		<h2 class="font-semibold">Über</h2>
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
