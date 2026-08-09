import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

export default defineConfig({
  plugins: [react()],
  server: {
    port: 5173,
    proxy: {
      '/auth': 'http://localhost:5260',
      '/users': 'http://localhost:5260',
      '/settings': 'http://localhost:5260'
    }
  }
});
