import { defineConfig } from 'vite';
import fs from 'fs';
import path from 'path';
import os from 'os';

export default defineConfig(({ command }) => {
  const isDev = command === 'serve';
  const certPath = path.join(os.homedir(), '.office-addin-dev-certs');
  
  return {
    base: '/rcc-addin/',
    server: {
      port: 5173,
      // Solo intenta leer los certificados si estamos corriendo `npm run dev` en local
      https: isDev ? {
        key: fs.readFileSync(path.join(certPath, 'localhost.key')),
        cert: fs.readFileSync(path.join(certPath, 'localhost.crt')),
      } : false,
    },
    build: {
      rollupOptions: {
        input: {
          taskpane: 'index.html',
          // Descomenta la siguiente línea solo si llegas a crear un commands.html en la raíz
          commands: 'commands.html', 
        },
      },
    },
  };
});