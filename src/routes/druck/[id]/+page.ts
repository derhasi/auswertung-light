import { error } from '@sveltejs/kit';
import { ladeVeranstaltung } from '$lib/stores/veranstaltung.svelte';

export async function load({ params, url }) {
	try {
		return {
			store: await ladeVeranstaltung(Number(params.id)),
			dok: url.searchParams.get('dok') ?? 'ergebnis',
			klasse: url.searchParams.get('klasse') ?? 'alle',
			adressen: url.searchParams.get('adressen') === '1',
			training: url.searchParams.get('training') !== '0'
		};
	} catch (e) {
		error(404, e instanceof Error ? e.message : String(e));
	}
}
