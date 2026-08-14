import { defineConfig } from 'vite';
import basicSsl from '@vitejs/plugin-basic-ssl';
import { resolve } from 'path';

export default defineConfig({
  plugins: [basicSsl()],
  base: '/mailpilot/', // Important for GitHub Pages
  server: {
    port: 3000,
    https: true,
    host: 'localhost',
    headers: {
      'Access-Control-Allow-Origin': '*',
    },
  },
  build: {
    rollupOptions: {
      input: {
        taskpane: resolve(__dirname, 'src/taskpane.html'),
        privacy: resolve(__dirname, 'src/privacy.html'),
        terms: resolve(__dirname, 'src/terms.html'),
      },
    },
  },
});
