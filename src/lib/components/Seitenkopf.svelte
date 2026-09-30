<script lang="ts">
	import type { Snippet } from 'svelte';

	let {
		titel,
		untertitel,
		aktionen,
		mitte
	}: { titel: string; untertitel?: string; aktionen?: Snippet; mitte?: Snippet } = $props();
</script>

<header class="flex flex-wrap items-end gap-x-6 gap-y-3 border-b border-line bg-surface px-7 py-5 print:hidden">
	<div class="min-w-0">
		<h1 class="display text-[40px] leading-none">{titel}</h1>
		{#if untertitel}<p class="mt-1.5 text-base text-muted">{untertitel}</p>{/if}
	</div>
	{#if mitte}<div class="flex flex-wrap items-center gap-1.5 self-center">{@render mitte()}</div>{/if}
	{#if aktionen}<div class="ml-auto flex flex-wrap gap-2">{@render aktionen()}</div>{/if}
</header>
