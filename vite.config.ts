import { sveltekit } from '@sveltejs/kit/vite';
import tailwindcss from '@tailwindcss/vite';
import { execSync } from 'node:child_process';
import { defineConfig } from 'vitest/config';

// Commit des Builds, damit auch Testversionen eindeutig zugeordnet werden können.
function commit(): string {
	try {
		const sha = execSync('git rev-parse HEAD', { encoding: 'utf8' }).trim();
		const geaendert = execSync('git status --porcelain --untracked-files=no', { encoding: 'utf8' }).trim() !== '';
		return geaendert ? `${sha}-geändert` : sha;
	} catch {
		return process.env.GITHUB_SHA ?? '';
	}
}

const host = process.env.TAURI_DEV_HOST;

export default defineConfig({
	plugins: [tailwindcss(), sveltekit()],
	define: {
		__COMMIT__: JSON.stringify(commit()),
		__BUILD_ZEIT__: JSON.stringify(new Date().toISOString())
	},
	clearScreen: false,
	server: {
		port: 1420,
		strictPort: true,
		host: host || false,
		hmr: host ? { protocol: 'ws', host, port: 1421 } : undefined,
		watch: { ignored: ['**/src-tauri/**'] }
	},
	build: {
		rolldownOptions: {
			// Der Hinweis zu Plugin-Laufzeiten betrifft SvelteKits eigenen Build-Schritt und ist nicht behebbar.
			checks: { bundlerTimings: false }
		}
	},
	test: {
		include: ['src/**/*.test.ts'],
		environment: 'node'
	}
});
