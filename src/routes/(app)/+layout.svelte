<script lang="ts">
	import { page } from '$app/state';
	import { CalendarDays, CircleHelp, Users } from '@lucide/svelte';
	import { VERSION } from '$lib/version';

	let { children } = $props();

	const navigation = [
		{ href: '/', titel: 'Veranstaltungen', icon: CalendarDays, aktiv: (p: string) => p === '/' || p.startsWith('/veranstaltung') },
		{ href: '/fahrer', titel: 'Fahrerdatenbank', icon: Users, aktiv: (p: string) => p.startsWith('/fahrer') },
		{ href: '/info', titel: 'Hilfe & Info', icon: CircleHelp, aktiv: (p: string) => p.startsWith('/info') }
	];
</script>

<div class="flex h-screen overflow-hidden">
	<nav class="flex w-56 shrink-0 flex-col bg-nav text-nav-fg print:hidden">
		<a href="/" class="flex items-center gap-2.5 px-5 py-5">
			<img src="/favicon.png" alt="" class="size-8 rounded-lg" />
			<div class="leading-tight">
				<div class="text-sm font-bold text-white">Auswertung Light</div>
				<div class="text-[11px] opacity-70">Kart-Slalom · v{VERSION}</div>
			</div>
		</a>
		<ul class="flex flex-col gap-1 px-3">
			{#each navigation as n (n.href)}
				{@const Icon = n.icon}
				{@const aktiv = n.aktiv(page.url.pathname)}
				<li>
					<a
						href={n.href}
						class="flex items-center gap-3 rounded-lg px-3 py-2 text-sm font-medium transition-colors {aktiv
							? 'bg-white/10 text-white'
							: 'hover:bg-white/5 hover:text-white'}"
						aria-current={aktiv ? 'page' : undefined}
					>
						<Icon size={18} />
						{n.titel}
					</a>
				</li>
			{/each}
		</ul>
		<div class="mt-auto px-5 py-4 text-[11px] leading-relaxed opacity-60">
			Für Kart-Slalom-Vereine<br />zugspitzpokal.de
		</div>
	</nav>
	<main class="min-w-0 flex-1 overflow-y-auto">
		{@render children()}
	</main>
</div>
