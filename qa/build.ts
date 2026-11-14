/**
 * Builds a Slidev deck with slidev2pptx's source, writing its problems as JSON.
 * Runs in the deck's directory as its own process, like the CLI would.
 * @example
 * ```shell
 * node ../../../qa/build.ts slides.md out.pptx problems.json [--with-clicks]
 * ```
 */
import fs from "node:fs/promises";

import { buildPptx } from "../src/index.ts";

const [entry, output, problemsPath, ...flags] = process.argv.slice(2);
if (!entry || !output || !problemsPath) {
	throw new Error(
		"Usage: node qa/build.ts <entry> <output.pptx> <problems.json> [--with-clicks]",
	);
}

const started = performance.now();
const result = await buildPptx({
	entry,
	output,
	withClicks: flags.includes("--with-clicks"),
});
await fs.writeFile(
	problemsPath,
	JSON.stringify(
		{ ...result, seconds: (performance.now() - started) / 1000 },
		null,
		"\t",
	) + "\n",
);
