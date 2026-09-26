import { error } from '@sveltejs/kit';
import { ladeVeranstaltung } from '$lib/stores/veranstaltung.svelte';

export async function load({ params }) {
	try {
		return { store: await ladeVeranstaltung(Number(params.id)) };
	} catch (e) {
		error(404, e instanceof Error ? e.message : String(e));
	}
}
