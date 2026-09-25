<script lang="ts">
	import { CircleAlert, CircleCheck, Info, TriangleAlert, X } from '@lucide/svelte';
	import { ui } from './ui-zustand.svelte';
	import Dialog from './Dialog.svelte';

	const stil = {
		erfolg: 'border-ok/40 bg-ok-soft text-ok',
		fehler: 'border-danger/40 bg-danger-soft text-danger',
		info: 'border-info/40 bg-info-soft text-info',
		warnung: 'border-warn/40 bg-warn-soft text-warn'
	};
	const icons = { erfolg: CircleCheck, fehler: CircleAlert, info: Info, warnung: TriangleAlert };
</script>

<div class="pointer-events-none fixed right-4 bottom-4 z-50 flex w-96 max-w-[calc(100vw-2rem)] flex-col gap-2 print:hidden" aria-live="polite">
	{#each ui.meldungen as m (m.id)}
		{@const Icon = icons[m.art]}
		<div class="pointer-events-auto flex gap-3 rounded-xl border px-4 py-3 text-sm shadow-lg {stil[m.art]}" role="status">
			<Icon size={18} class="mt-0.5 shrink-0" />
			<div class="min-w-0 flex-1">
				<p class="font-medium">{m.text}</p>
				{#if m.details?.length}
					<ul class="mt-1 max-h-32 list-disc overflow-y-auto pl-4 text-xs opacity-90">
						{#each m.details as d, i (i)}<li>{d}</li>{/each}
					</ul>
				{/if}
			</div>
			<button class="shrink-0 opacity-60 hover:opacity-100" onclick={() => ui.schliessen(m.id)} aria-label="Meldung schließen"><X size={16} /></button>
		</div>
	{/each}
</div>

{#if ui.bestaetigung}
	{@const b = ui.bestaetigung}
	<Dialog offen={true} titel={b.titel} breite="max-w-md" onschliessen={() => b.antwort(false)}>
		<p class="text-sm whitespace-pre-line">{b.text}</p>
		{#snippet aktionen()}
			<button class="btn" onclick={() => b.antwort(false)}>{b.nein}</button>
			<!-- svelte-ignore a11y_autofocus -->
			<button class={b.gefaehrlich ? 'btn btn-danger' : 'btn btn-primary'} autofocus onclick={() => b.antwort(true)}>{b.ja}</button>
		{/snippet}
	</Dialog>
{/if}
