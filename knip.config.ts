import type { KnipConfig } from "knip";

export default {
	entry: ["qa/*.ts", "qa/*/*.ts", "src/ssr/*", "src/**/*.test.*"],
	ignore: ["src/fixture/**"],
	ignoreExportsUsedInFile: { interface: true, type: true },
	ignoreUnresolved: [/^#slidev\//],
	project: ["qa/**/*.ts", "src/**/*.{ts,vue}"],
	treatConfigHintsAsErrors: true,
} satisfies KnipConfig;
