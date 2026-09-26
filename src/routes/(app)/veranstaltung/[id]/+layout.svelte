<script lang="ts">
	import { page } from '$app/state';
	import { ClipboardList, Flag, LayoutDashboard, Printer, Settings, Timer, Trophy, Users } from '@lucide/svelte';
	import { formatDatum } from '$lib/domain/fahrer-import';

	let { data, children } = $props();
	const s = $derived(data.store);

	const tabs = [
		{ pfad: '', titel: 'Übersicht', icon: LayoutDashboard },
		{ pfad: '/nennung', titel: 'Nennung', icon: ClipboardList },
		{ pfad: '/erfassung', titel: 'Erfassung', icon: Timer },
		{ pfad: '/ergebnisse', titel: 'Ergebnisse', icon: Trophy },
		{ pfad: '/mannschaft', titel: 'Mannschaft', icon: Users },
		{ pfad: '/abschluss', titel: 'Drucken & Export', icon: Printer },
		{ pfad: '/einstellungen', titel: 'Einstellungen', icon: Settings }
	];
	const basis = $derived(`/veranstaltung/${s.id}`);
</script>

<div class="sticky top-0 z-20 border-b border-line bg-surface/95 backdrop-blur print:hidden">
	<div class="flex items-center gap-3 px-8 pt-5">
		<Flag size={20} class="text-accent" />
		<div class="min-w-0">
			<h1 class="truncate text-lg leading-tight font-bold">{s.v.name}</h1>
			<p class="text-xs text-muted">
				{formatDatum(s.v.datum)}{s.v.ort ? ` · ${s.v.ort}` : ''} · {s.starter.length} Starter
			</p>
		</div>
	</div>
	<nav class="mt-3 flex gap-1 overflow-x-auto px-6">
		{#each tabs as t (t.pfad)}
			{@const Icon = t.icon}
			{@const href = basis + t.pfad}
			{@const aktiv = page.url.pathname === href}
			<a
				{href}
				class="flex items-center gap-2 border-b-2 px-3 py-2.5 text-sm font-medium whitespace-nowrap transition-colors {aktiv
					? 'border-accent text-fg'
					: 'border-transparent text-muted hover:text-fg'}"
				aria-current={aktiv ? 'page' : undefined}
			>
				<Icon size={16} />
				{t.titel}
			</a>
		{/each}
	</nav>
</div>

{#key s.id}
	{@render children()}
{/key}
