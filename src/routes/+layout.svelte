<script lang="ts">
	import '../app.css';
	import Meldungen from '$lib/ui/Meldungen.svelte';

	let { children } = $props();
</script>

{@render children()}
<Meldungen />
