import alias from '@rollup/plugin-alias';
import typescript from '@rollup/plugin-typescript';
import nodeExternals from 'rollup-plugin-node-externals';
import dts from 'rollup-plugin-dts';
import { fileURLToPath } from 'node:url';

const input = 'src/docx-preview.ts';

const aliasPlugin = alias({
	entries: [{ find: '@docx', replacement: fileURLToPath(new URL('./src', import.meta.url)) }],
});

const umdGlobals = {
	jszip: 'JSZip',
	konva: 'Konva',
	'lodash-es': '_',
};

export default [
	{
		input,
		output: [
			{
				file: 'dist/docx-renderer.mjs',
				format: 'es',
				sourcemap: true,
			},
			{
				file: 'dist/docx-renderer.cjs',
				format: 'cjs',
				exports: 'named',
				sourcemap: true,
			},
			{
				// Browser/global build used by the Playwright test harness.
				file: 'dist/docx-renderer.umd.js',
				format: 'umd',
				name: 'docx',
				globals: umdGlobals,
				sourcemap: true,
			},
			{
				// Same UMD build, served by the demo pages under docs/.
				file: 'docs/js/docx-renderer.js',
				format: 'umd',
				name: 'docx',
				globals: umdGlobals,
				sourcemap: true,
			},
		],
		plugins: [
			aliasPlugin,
			nodeExternals(),
			typescript(),
		],
	},
	{
		input,
		output: {
			file: 'dist/docx-renderer.d.ts',
			format: 'es',
		},
		plugins: [
			aliasPlugin,
			dts(),
		],
	},
];
