import { defineConfig } from 'vite';
import { fileURLToPath, URL } from 'node:url';

export default defineConfig({
	root: '.',
	resolve: {
		alias: {
			'@docx': fileURLToPath(new URL('./src', import.meta.url)),
		},
	},
	server: {
		open: '/playground/index.html',
	},
});
