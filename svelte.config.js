import adapter from '@sveltejs/adapter-static';
import { vitePreprocess } from '@sveltejs/vite-plugin-svelte';

/** @type {import('@sveltejs/kit').Config} */
const config = {
	preprocess: vitePreprocess(),
	kit: {
		// Tauri liefert eine Single-Page-App aus: alle Routen fallen auf index.html zurück.
		adapter: adapter({ fallback: 'index.html' })
	}
};

export default config;
