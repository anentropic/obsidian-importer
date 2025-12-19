import path from 'path';
import { defineConfig } from 'vitest/config';

export default defineConfig({
	test: {
		include: ['src/**/*.test.ts', 'tests/**/*.test.ts'],
		environment: 'node',
		alias: {
			obsidian: path.resolve(__dirname, 'tests/mocks/obsidian.ts'),
		},
	},
});
