import { defineConfig } from 'vite';
import { fileURLToPath, URL } from 'node:url';

export default defineConfig({
	root: 'benchmark',
	resolve: {
		alias: {
			'@docx': fileURLToPath(new URL('../src', import.meta.url)),
		},
	},
	server: {
		fs: {
			allow: ['..'],
		},
	},
});
