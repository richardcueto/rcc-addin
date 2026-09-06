import react from '@vitejs/plugin-react'
import { defineConfig } from 'vite'
import fs from 'fs';
import path from 'path';
import os from 'os';

const certPath = path.join(os.homedir(), '.office-addin-dev-certs');

// https://vite.dev/config/
export default defineConfig({
  plugins: [
    react()
  ],
  base: '/rcc-addin/',
  server: {
    port: 5173,
    https: {
      key: fs.readFileSync(path.join(certPath, 'localhost.key')),
      cert: fs.readFileSync(path.join(certPath, 'localhost.crt')),
    },
  },
})
