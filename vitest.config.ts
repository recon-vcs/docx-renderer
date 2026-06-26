import { defineConfig } from 'vitest/config';
import { fileURLToPath, URL } from 'node:url';

export default defineConfig({
	resolve: {
		alias: {
			'@docx': fileURLToPath(new URL('./src', import.meta.url)),
		},
	},
	test: {
		// jsdom provides DOMParser/XML DOM used by the parsing pipeline.
		environment: 'jsdom',
		include: ['tests/unit/**/*.spec.ts'],
	},
});
