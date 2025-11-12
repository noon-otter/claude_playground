import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import fs from 'fs';
import path from 'path';

// https://vitejs.dev/config/
export default defineConfig({
  plugins: [react()],
  root: '.',
  build: {
    outDir: 'dist',
    rollupOptions: {
      input: {
        main: path.resolve(__dirname, 'index.html'),
        commands: path.resolve(__dirname, 'commands.html'),
        register: path.resolve(__dirname, 'register.html')
      }
    }
  },
  server: {
    port: 3000,
    https: {
      key: fs.existsSync('./certs/localhost.key')
        ? fs.readFileSync('./certs/localhost.key')
        : undefined,
      cert: fs.existsSync('./certs/localhost.crt')
        ? fs.readFileSync('./certs/localhost.crt')
        : undefined
    },
    cors: true
  },
  resolve: {
    alias: {
      '@': path.resolve(__dirname, './src')
    }
  }
});
