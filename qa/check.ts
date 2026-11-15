/**
 * Builds one corpus deck from the working tree into another directory, renders
 * it, and compares states against the deck's last run, leaving the run as it
 * is: a quick check of a fix while a batch uses the runs.
 * Compares the given states, or every state that didn't pass.
 * @example
 * ```shell
 * node qa/check.ts 2026-05-21-what-s-new-in-typescript-6-0 /tmp/check 021-01
 * ```
 */
import { execFile } from "node:child_process";
import fs from "node:fs/promises";
import path from "node:path";
import { promisify } from "node:util";

import { compareImages } from "./compare/images.ts";
import { renderWithPowerPoint } from "./render/powerpoint.ts";

const run = promisify(execFile);

const [id, output, ...names] = process.argv.slice(2);
if (!id || !output) {
	throw new Error(
		"Usage: node qa/check.ts <deck id> <output directory> [states...]",
	);
}

const deck = path.join(import.meta.dirname, "runs", id);
const pptx = path.join(output, `${id}.pptx`);
await fs.mkdir(output, { recursive: true });
await run(
	process.execPath,
	[
		path.join(import.meta.dirname, "build.ts"),
		"slides.md",
		pptx,
		path.join(output, `${id}.json`),
		"--with-clicks",
	],
	{ cwd: path.join(deck, "slidev"), maxBuffer: 64 * 1024 * 1024 },
);

const pngs = path.join(output, `${id}-png`);
await fs.rm(pngs, { force: true, recursive: true });
await renderWithPowerPoint(pptx, pngs);
const files = (await fs.readdir(pngs))
	.filter((file) => file.endsWith(".png"))
	.sort();

const report = JSON.parse(
	await fs.readFile(path.join(deck, "report.json"), "utf8"),
) as {
	states: {
		name: string;
		step2?: { ssim: number };
		verdict: string;
	}[];
};
for (const [index, state] of report.states.entries()) {
	if (names.length ? !names.includes(state.name) : state.verdict === "pass") {
		continue;
	}

	const now = await compareImages(
		path.join(deck, "slidev-png", `${state.name}.png`),
		path.join(pngs, files[index]),
		path.join(output, `${id}-${state.name}-diff.png`),
	);
	console.log(
		`${state.name}: SSIM ${String(state.step2?.ssim.toFixed(4))} → ${now.ssim.toFixed(4)}, ${now.changed.toFixed(4)} changed`,
	);
}
