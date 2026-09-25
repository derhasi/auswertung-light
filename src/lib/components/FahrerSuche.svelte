<script lang="ts">
	import { Search } from '@lucide/svelte';
	import type { Fahrer } from '$lib/db';

	interface Props {
		fahrer: Fahrer[];
		gemeldet: Set<string>;
		onauswahl: (f: Fahrer) => void;
	}

	let { fahrer, gemeldet, onauswahl }: Props = $props();
	let text = $state('');
	let markiert = $state(0);
	let offen = $state(false);
	let eingabe: HTMLInputElement | undefined = $state();

	const treffer = $derived.by(() => {
		const begriffe = text.toLocaleLowerCase('de').split(/\s+/).filter(Boolean);
		if (!begriffe.length) return [];
		return fahrer
			.filter((f) => begriffe.every((b) => `${f.lizenz} ${f.nachname} ${f.vorname} ${f.verein}`.toLocaleLowerCase('de').includes(b)))
			.slice(0, 8);
	});

	export function fokussieren() {
		eingabe?.focus();
	}

	function waehlen(f: Fahrer | undefined) {
		if (!f) return;
		onauswahl(f);
		text = '';
		offen = false;
		markiert = 0;
	}

	function taste(e: KeyboardEvent) {
		if (e.key === 'ArrowDown') {
			e.preventDefault();
			markiert = Math.min(markiert + 1, treffer.length - 1);
		} else if (e.key === 'ArrowUp') {
			e.preventDefault();
			markiert = Math.max(markiert - 1, 0);
		} else if (e.key === 'Enter') {
			e.preventDefault();
			waehlen(treffer[markiert]);
		} else if (e.key === 'Escape') {
			text = '';
		}
	}
</script>

<div class="relative">
	<Search size={18} class="pointer-events-none absolute top-1/2 left-3.5 -translate-y-1/2 text-muted" />
	<input
		bind:this={eingabe}
		class="input py-3 pl-11 text-base"
		placeholder="Fahrer suchen: Name, Lizenz oder Verein …"
		bind:value={text}
		oninput={() => {
			offen = true;
			markiert = 0;
		}}
		onfocus={() => (offen = true)}
		onblur={() => setTimeout(() => (offen = false), 150)}
		onkeydown={taste}
		role="combobox"
		aria-expanded={offen && treffer.length > 0}
		aria-controls="fahrer-treffer"
		aria-autocomplete="list"
	/>
	{#if offen && text.trim()}
		<ul id="fahrer-treffer" class="absolute z-30 mt-1 w-full overflow-hidden rounded-xl border border-line bg-surface shadow-xl" role="listbox">
			{#each treffer as f, i (f.id)}
				<li role="option" aria-selected={i === markiert}>
					<button
						class="flex w-full items-center gap-3 px-4 py-2.5 text-left text-sm {i === markiert ? 'bg-accent-soft' : 'hover:bg-sunken'}"
						onmousedown={(e) => e.preventDefault()}
						onclick={() => waehlen(f)}
						onmouseenter={() => (markiert = i)}
					>
						<span class="w-16 font-mono text-xs text-muted">{f.lizenz}</span>
						<span class="flex-1 font-medium">{f.nachname}, {f.vorname}</span>
						<span class="text-muted">{f.verein}</span>
						<span class="badge bg-sunken text-muted">{f.klasse || '–'}</span>
						{#if gemeldet.has(f.lizenz)}<span class="badge bg-warn-soft text-warn">bereits gemeldet</span>{/if}
					</button>
				</li>
			{:else}
				<li class="px-4 py-3 text-sm text-muted">Kein Fahrer gefunden.</li>
			{/each}
		</ul>
	{/if}
</div>
