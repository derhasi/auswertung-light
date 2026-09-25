<script lang="ts">
	import { X } from '@lucide/svelte';
	import type { Snippet } from 'svelte';

	interface Props {
		offen: boolean;
		titel: string;
		breite?: string;
		onschliessen: () => void;
		children: Snippet;
		aktionen?: Snippet;
	}

	let { offen, titel, breite = 'max-w-lg', onschliessen, children, aktionen }: Props = $props();
	let dialog: HTMLDialogElement | undefined = $state();

	$effect(() => {
		if (!dialog) return;
		if (offen && !dialog.open) dialog.showModal();
		if (!offen && dialog.open) dialog.close();
	});
</script>

<dialog
	bind:this={dialog}
	class="m-auto w-[calc(100%-2rem)] {breite} rounded-2xl border border-line bg-surface p-0 text-fg shadow-2xl backdrop:bg-black/40 backdrop:backdrop-blur-[2px]"
	onclose={() => offen && onschliessen()}
	oncancel={(e) => {
		e.preventDefault();
		onschliessen();
	}}
>
	{#if offen}
		<div class="flex items-center justify-between border-b border-line px-5 py-3.5">
			<h2 class="text-base font-semibold">{titel}</h2>
			<button class="btn btn-ghost btn-icon -mr-2" onclick={onschliessen} aria-label="Schließen"><X size={18} /></button>
		</div>
		<div class="max-h-[70vh] overflow-y-auto px-5 py-4">
			{@render children()}
		</div>
		{#if aktionen}
			<div class="flex justify-end gap-2 border-t border-line bg-sunken/50 px-5 py-3">
				{@render aktionen()}
			</div>
		{/if}
	{/if}
</dialog>
