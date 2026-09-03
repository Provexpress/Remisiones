import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const projectRoot = fileURLToPath(new URL('.', import.meta.url));

export default defineConfig({
  plugins: [react()],
  build: {
    target: 'es2022',
    sourcemap: true,
    rolldownOptions: {
      input: {
        app: resolve(projectRoot, 'index.html'),
        redirect: resolve(projectRoot, 'redirect.html'),
      },
    },
  },
});
