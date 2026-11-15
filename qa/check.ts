/**
 * Builds one corpus deck from the working tree into another directory, renders
 * the given states (or every state that didn't pass), and compares them
 * against the deck's last run, leaving the run as it is: a quick check of a
 * fix while a batch uses the runs.
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
import { type RenderedState, rerenderMisses } from "./compare/rerender.ts";
import { renderSlidesWithPowerPoint } from "./render/powerpoint.ts";

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

const report = JSON.parse(
	await fs.readFile(path.join(deck, "report.json"), "utf8"),
) as {
	states: {
		name: string;
		step2?: { ssim: number };
		verdict: string;
	}[];
};
const checked = report.states.flatMap((state, index) =>
	(names.length ? names.includes(state.name) : state.verdict !== "pass")
		? [{ ...state, index }]
		: [],
);

const pngs = path.join(output, `${id}-png`);
const files = await renderSlidesWithPowerPoint(
	pptx,
	checked.map((state) => state.index),
	pngs,
);
const rendered: RenderedState[] = [];
for (const [position, state] of checked.entries()) {
	const expected = path.join(deck, "slidev-png", `${state.name}.png`);
	const diff = path.join(output, `${id}-${state.name}-diff.png`);
	rendered.push({
		actual: files[position],
		comparison: await compareImages(expected, files[position], diff),
		diff,
		expected,
		index: state.index,
		render: 1,
	});
}

await rerenderMisses(pptx, rendered, path.join(output, `${id}-again`));

for (const [position, state] of checked.entries()) {
	const now = rendered[position];
	console.log(
		`${state.name}: SSIM ${String(state.step2?.ssim.toFixed(4))} → ${now.comparison.ssim.toFixed(4)}, ${now.comparison.changed.toFixed(4)} changed (render ${String(now.render)})`,
	);
}
