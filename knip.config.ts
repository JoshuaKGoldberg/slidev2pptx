import type { KnipConfig } from "knip";

export default {
	entry: ["src/ssr/*", "src/**/*.test.*"],
	ignore: ["src/fixture/**"],
	ignoreExportsUsedInFile: { interface: true, type: true },
	ignoreUnresolved: [/^#slidev\//],
	project: ["src/**/*.{ts,vue}"],
	treatConfigHintsAsErrors: true,
} satisfies KnipConfig;
