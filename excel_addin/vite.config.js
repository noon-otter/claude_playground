import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import fs from 'fs';
import path from 'path';
import { homedir } from 'os';

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
      key: fs.readFileSync(path.join(homedir(), '.office-addin-dev-certs', 'localhost.key')),
      cert: fs.readFileSync(path.join(homedir(), '.office-addin-dev-certs', 'localhost.crt'))
    },
    cors: true
  },
  resolve: {
    alias: {
      '@': path.resolve(__dirname, './src')
    }
  }
});
