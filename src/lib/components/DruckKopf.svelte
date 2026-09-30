<script lang="ts">
	import type { Veranstaltung } from '$lib/db';
	import { formatDatum } from '$lib/domain/fahrer-import';

	let { v, titel, untertitel = '' }: { v: Veranstaltung; titel: string; untertitel?: string } = $props();
</script>

<header class="mb-4 flex items-center gap-4 border-b-2 border-fg pb-3">
	<div class="flex h-16 w-24 items-center justify-start">
		{#if v.logoLinks}<img src={v.logoLinks} alt="" class="max-h-16 max-w-24 object-contain" />{/if}
	</div>
	<div class="flex-1 text-center">
		<p class="text-[10pt] text-muted">{v.name} · {formatDatum(v.datum)}{v.ort ? ` · ${v.ort}` : ''}</p>
		<h1 class="text-[18pt] leading-tight font-bold">{titel}</h1>
		{#if untertitel}<p class="text-[11pt] font-semibold">{untertitel}</p>{/if}
		{#if v.ausrichter}<p class="text-[9pt] text-muted">Ausrichter: {v.ausrichter}</p>{/if}
	</div>
	<div class="flex h-16 w-24 items-center justify-end">
		{#if v.logoRechts}<img src={v.logoRechts} alt="" class="max-h-16 max-w-24 object-contain" />{/if}
	</div>
</header>
