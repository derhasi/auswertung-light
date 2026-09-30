<script lang="ts">
	let { data, children } = $props();
</script>

{#key data.store.id}
	{@render children()}
{/key}
