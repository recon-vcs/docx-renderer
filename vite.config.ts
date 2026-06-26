import { defineConfig } from 'vite';
import { fileURLToPath, URL } from 'node:url';

export default defineConfig({
	root: 'playground',
	resolve: {
		alias: {
			'@docx': fileURLToPath(new URL('./src', import.meta.url)),
		},
	},
	server: {
		open: '/',
	},
	build: {
		outDir: '../dist-playground',
		emptyOutDir: true,
	},
});
