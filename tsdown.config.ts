import { defineConfig } from "tsdown";

export default defineConfig({
	// Vite loads these stand-ins from source while building a deck, so they ship next to render.mjs.
	copy: [{ from: "src/ssr/*", to: "dist/ssr" }],
	entry: ["src/**/*.ts", "!src/**/*.test.*", "!src/fixture/**", "!src/ssr/**"],
	unbundle: true,
});
