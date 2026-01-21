import { defineConfig } from "vite";
import react from "@vitejs/plugin-react-swc";
import path from "path";
import { componentTagger } from "lovable-tagger";

// https://vitejs.dev/config/
export default defineConfig(({ mode }) => ({
  server: {
    host: "::",
    port: 8080,
  },
  plugins: [react(), mode === "development" && componentTagger()].filter(Boolean),
  resolve: {
    alias: {
      "@": path.resolve(__dirname, "./src"),
    },
    // Deduplicate React to fix "Cannot read properties of null (reading 'useState')"
    // This ensures the SharePoint SDK uses the same React instance as the app
    dedupe: ["react", "react-dom"],
  },
  optimizeDeps: {
    // Force pre-bundling of these to ensure single instance
    include: ["react", "react-dom"],
  },
}));
