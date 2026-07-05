import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// En dev, el frontend corre en :5173 y el backend en :3000.
// Se proxean las llamadas /api al server Express.
export default defineConfig({
  plugins: [react()],
  server: {
    proxy: {
      "/api": "http://localhost:3000",
    },
  },
  build: {
    outDir: "dist",
  },
});
