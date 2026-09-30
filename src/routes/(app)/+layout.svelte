<script lang="ts">
	import { page } from '$app/state';
	import { CircleHelp, Users } from '@lucide/svelte';
	import { VERSION } from '$lib/version';
	import { formatDatum } from '$lib/domain/fahrer-import';
	import type { VeranstaltungsStore } from '$lib/stores/veranstaltung.svelte';

	let { children } = $props();

	const navigation = [
		{ href: '/', titel: 'Veranstaltungen', aktiv: (p: string) => p === '/' || p.startsWith('/veranstaltung') },
		{ href: '/fahrer', titel: 'Fahrerdatenbank', aktiv: (p: string) => p.startsWith('/fahrer') },
		{ href: '/info', titel: 'Hilfe & Info', aktiv: (p: string) => p.startsWith('/info') }
	];

	const tabs = [
		{ pfad: '', titel: 'Übersicht' },
		{ pfad: '/nennung', titel: 'Nennung' },
		{ pfad: '/erfassung', titel: 'Erfassung' },
		{ pfad: '/ergebnisse', titel: 'Ergebnisse' },
		{ pfad: '/mannschaft', titel: 'Mannschaft' },
		{ pfad: '/abschluss', titel: 'Drucken' },
		{ pfad: '/einstellungen', titel: 'Einstellungen' }
	];

	/** Innerhalb einer Veranstaltung zeigt der Kopf deren Reiter statt der Hauptnavigation. */
	const store = $derived(page.data.store as VeranstaltungsStore | undefined);
	const basis = $derived(store ? `/veranstaltung/${store.id}` : '');
</script>

<div class="flex h-screen flex-col overflow-hidden">
	<header class="flex h-16 shrink-0 items-center gap-7 bg-nav px-7 text-white print:hidden">
		<a href="/" class="flex shrink-0 items-center gap-2.5" aria-label="Auswertung Light – Veranstaltungen">
			<svg width="26" height="26" viewBox="0 0 24 24" fill="none" stroke="var(--color-accent)" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
				<path d="M12 3 5 20h14L12 3z" /><path d="M8.5 12h7M7 16h10" />
			</svg>
			<span class="display text-2xl whitespace-nowrap {store ? 'hidden 2xl:inline' : ''}">Auswertung Light</span>
		</a>
		<nav aria-label={store ? 'Veranstaltung' : 'Hauptnavigation'} class="flex h-full min-w-0 gap-1 overflow-x-auto text-[15px] font-semibold">
			{#if store}
				{#each tabs as t (t.pfad)}
					{@const href = basis + t.pfad}
					{@const aktiv = page.url.pathname === href}
					<a
						{href}
						class="flex items-center border-b-[3px] px-3 pt-[3px] whitespace-nowrap transition-colors {aktiv
							? 'border-accent text-white'
							: 'border-transparent text-nav-fg hover:text-white'}"
						aria-current={aktiv ? 'page' : undefined}>{t.titel}</a
					>
				{/each}
			{:else}
				{#each navigation as n (n.href)}
					{@const aktiv = n.aktiv(page.url.pathname)}
					<a
						href={n.href}
						class="flex items-center border-b-[3px] px-3 pt-[3px] whitespace-nowrap transition-colors {aktiv
							? 'border-accent text-white'
							: 'border-transparent text-nav-fg hover:text-white'}"
						aria-current={aktiv ? 'page' : undefined}>{n.titel}</a
					>
				{/each}
			{/if}
		</nav>
		<div class="ml-auto flex shrink-0 items-center gap-4">
			{#if store}
				<div class="hidden max-w-72 text-right leading-tight xl:block">
					<div class="truncate text-[15px] font-bold">{store.v.name}</div>
					<div class="truncate text-[13px] text-nav-fg">{formatDatum(store.v.datum)}{store.v.ort ? ` · ${store.v.ort}` : ''}</div>
				</div>
				<div class="flex gap-1 xl:border-l xl:border-ink-line xl:pl-3">
					<a href="/fahrer" class="rounded-lg p-2 text-nav-fg hover:bg-white/10 hover:text-white" aria-label="Fahrerdatenbank" title="Fahrerdatenbank"><Users size={20} /></a>
					<a href="/info" class="rounded-lg p-2 text-nav-fg hover:bg-white/10 hover:text-white" aria-label="Hilfe & Info" title="Hilfe & Info"><CircleHelp size={20} /></a>
				</div>
			{:else}
				<span class="text-[13px] text-nav-fg">Kart-Slalom · v{VERSION}</span>
			{/if}
		</div>
	</header>
	<main class="min-h-0 min-w-0 flex-1 overflow-y-auto">
		{@render children()}
	</main>
</div>
