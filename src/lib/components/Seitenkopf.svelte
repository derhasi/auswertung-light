<script lang="ts">
	import type { Snippet } from 'svelte';

	let { titel, untertitel, aktionen }: { titel: string; untertitel?: string; aktionen?: Snippet } = $props();
</script>

<header class="flex flex-wrap items-end justify-between gap-4 px-8 pt-8 pb-6">
	<div>
		<h1 class="text-2xl font-bold tracking-tight">{titel}</h1>
		{#if untertitel}<p class="mt-1 text-sm text-muted">{untertitel}</p>{/if}
	</div>
	{#if aktionen}<div class="flex flex-wrap gap-2">{@render aktionen()}</div>{/if}
</header>
