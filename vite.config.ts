import { defineConfig } from 'vitest/config';
import react from '@vitejs/plugin-react';
import tailwindcss from '@tailwindcss/vite';
import { viteSingleFile } from 'vite-plugin-singlefile';

// Genera un único index.html autocontenido: se puede enviar por correo/WhatsApp
// y abrir sin internet con doble clic.
export default defineConfig({
  plugins: [react(), tailwindcss(), viteSingleFile()],
  build: { chunkSizeWarningLimit: 4000 },
  test: { testTimeout: 60000 },
});
