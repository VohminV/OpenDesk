import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

// https://vitejs.dev/config/
// base './' — относительные пути, чтобы работало и на GitHub Pages
// (https://user.github.io/repo/), и на своём домене, и из file://.
export default defineConfig({
  base: './',
  plugins: [react()],
  server: {
    port: 5173,
  },
  build: {
    chunkSizeWarningLimit: 1000,
    rollupOptions: {
      output: {
        manualChunks: {
          'vendor-react': ['react', 'react-dom'],
          'vendor-docx': ['docx', 'mammoth'],
          'vendor-exceljs': ['exceljs'],
        },
      },
    },
  },
});
