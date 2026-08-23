import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// Using a relative base so the built assets work correctly
// when served from a GitHub Pages project path (https://<user>.github.io/<repo>/)
export default defineConfig({
  plugins: [react()],
  base: "./",
});
