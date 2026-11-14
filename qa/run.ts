/**
 * Round-trips one corpus deck: pptx → Slidev (converter) → pptx (slidev2pptx),
 * rendering each and comparing the Slidev rendering against slidev2pptx's.
 * Steps whose outputs exist are skipped unless --from names an earlier step.
 * @example
 * ```shell
 * node qa/run.ts 2026-05-21-what-s-new-in-typescript-6-0 --from build
 * ```
 */
import { execFile } from "node:child_process";
import fs from "node:fs/promises";
import path from "node:path";
import { parseArgs, promisify } from "node:util";

import type { BuildPptxResult } from "../src/types.ts";
import type { ConversionLog } from "./pptx2slidev/convert.ts";

import { compareImages } from "./compare/images.ts";
import { deckPath } from "./corpus/paths.ts";
import { renderWithPowerPoint } from "./render/powerpoint.ts";
import { renderWithSlidev } from "./render/slidev.ts";

const run = promisify(execFile);

const STEPS = ["convert", "slidev", "build", "render", "compare"] as const;

/**
 * Thresholds on the normalized comparison, calibrated on states that look the
 * same (about 0.4% changed, 0.991 SSIM) and that visibly differ (over 1%).
 * They rank what reviewers look at; reviewers decide what's a real difference.
 */
export const THRESHOLDS = {
	changed: { fail: 0.03, review: 0.008 },
	ssim: { fail: 0.97, review: 0.985 },
};

const { positionals, values } = parseArgs({
	allowPositionals: true,
	options: { from: { type: "string" } },
});
const [id] = positionals;
if (!id) {
	throw new Error("Usage: node qa/run.ts <deck id> [--from <step>]");
}

const from = values.from
	? STEPS.indexOf(values.from as (typeof STEPS)[number])
	: STEPS.length;
const directory = path.join(import.meta.dirname, "runs", id);
const paths = {
	animated: path.join(directory, "pptx/animated.pptx"),
	animatedProblems: path.join(directory, "pptx/animated.json"),
	clicks: path.join(directory, "pptx/clicks.pptx"),
	clicksProblems: path.join(directory, "pptx/clicks.json"),
	diff: path.join(directory, "diff"),
	entry: path.join(directory, "slidev/slides.md"),
	log: path.join(directory, "slidev/conversion-log.json"),
	originalPng: path.join(directory, "original-png"),
	pptxPng: path.join(directory, "pptx-png"),
	report: path.join(directory, "report.json"),
	sheets: path.join(directory, "sheets"),
	slidevErrors: path.join(directory, "slidev-errors.json"),
	slidevPng: path.join(directory, "slidev-png"),
};

const timings: Record<string, number> = {};

await step("convert", paths.entry, async () => {
	await run(process.execPath, [
		path.join(import.meta.dirname, "pptx2slidev/convert.ts"),
		deckPath(id),
		path.dirname(paths.entry),
	]);
});

await step("slidev", paths.slidevPng, async () => {
	const { errors } = await renderWithSlidev(paths.entry, paths.slidevPng);
	await fs.writeFile(
		paths.slidevErrors,
		JSON.stringify(errors, null, "\t") + "\n",
	);
});

await step("build", paths.clicks, async () => {
	await fs.mkdir(path.dirname(paths.clicks), { recursive: true });
	const build = (output: string, problems: string, flags: string[]) =>
		run(
			process.execPath,
			[
				path.join(import.meta.dirname, "build.ts"),
				"slides.md",
				output,
				problems,
				...flags,
			],
			{ cwd: path.dirname(paths.entry), maxBuffer: 64 * 1024 * 1024 },
		);
	await build(paths.clicks, paths.clicksProblems, ["--with-clicks"]);
	await build(paths.animated, paths.animatedProblems, []);
});

await step("render", paths.pptxPng, async () => {
	await renderWithPowerPoint(paths.clicks, paths.pptxPng);
	// Hand-written repro decks have no original .pptx.
	if (!(await exists(paths.originalPng)) && (await exists(deckPath(id)))) {
		await renderWithPowerPoint(deckPath(id), paths.originalPng);
	}
});

await step("compare", paths.report, async () => {
	await fs.rm(paths.diff, { force: true, recursive: true });
	await fs.rm(paths.sheets, { force: true, recursive: true });
	await fs.mkdir(paths.diff, { recursive: true });
	await fs.mkdir(paths.sheets, { recursive: true });

	const slidev = await pngs(paths.slidevPng);
	const pptx = await pngs(paths.pptxPng);
	const original = (await exists(paths.originalPng))
		? await pngs(paths.originalPng)
		: [];
	const log = (await exists(paths.log))
		? (JSON.parse(await fs.readFile(paths.log, "utf8")) as ConversionLog)
		: { fonts: { missing: [] }, slides: [] };
	const clicks = JSON.parse(
		await fs.readFile(paths.clicksProblems, "utf8"),
	) as BuildPptxResult;
	const animated = JSON.parse(
		await fs.readFile(paths.animatedProblems, "utf8"),
	) as BuildPptxResult;
	const slidevErrors = JSON.parse(
		await fs.readFile(paths.slidevErrors, "utf8"),
	) as string[];

	// Each slide's last click state matches PowerPoint's export of the original, which shows every build.
	const lastStates = new Map<number, number>();
	for (const [index, file] of slidev.entries()) {
		lastStates.set(Number(path.basename(file).slice(0, 3)), index);
	}

	const states = [];
	for (const [index, file] of slidev.entries()) {
		const name = path.basename(file, ".png");
		const slide = Number(name.slice(0, 3));
		const click = Number(name.slice(4)) - 1;
		const actual = pptx.at(index);
		const comparison = actual
			? await compareImages(file, actual, path.join(paths.diff, `${name}.png`))
			: undefined;
		const originalFile =
			lastStates.get(slide) === index ? original.at(slide - 1) : undefined;
		const step1 = originalFile
			? await compareImages(
					originalFile,
					file,
					path.join(paths.diff, `${name}-original.png`),
				)
			: undefined;
		const verdict = !comparison
			? "missing"
			: comparison.ssim < THRESHOLDS.ssim.fail ||
				  comparison.changed > THRESHOLDS.changed.fail
				? "fail"
				: comparison.ssim < THRESHOLDS.ssim.review ||
					  comparison.changed > THRESHOLDS.changed.review
					? "review"
					: "pass";
		if (verdict !== "pass" && actual) {
			await run("magick", [
				"(",
				originalFile ?? file,
				file,
				"+append",
				")",
				"(",
				actual,
				path.join(paths.diff, `${name}.png`),
				"+append",
				")",
				"-append",
				"-resize",
				"1280x",
				path.join(paths.sheets, `${name}.png`),
			]);
		}

		states.push({
			click,
			name,
			original: step1,
			slide,
			step2: comparison,
			verdict,
		});
	}

	const report = {
		build: {
			animated: {
				problems: animated.problems,
				seconds: (animated as { seconds?: number }).seconds,
				slides: animated.slides,
			},
			clicks: {
				problems: clicks.problems,
				seconds: (clicks as { seconds?: number }).seconds,
				slides: clicks.slides,
			},
		},
		conversion: {
			losses: log.slides.reduce((sum, slide) => sum + slide.losses.length, 0),
			missingFonts: log.fonts.missing,
		},
		counts: {
			original: original.length,
			pptx: pptx.length,
			slidev: slidev.length,
		},
		id,
		slidevErrors,
		states,
		summary: {
			fail: states.filter((state) => state.verdict === "fail").length,
			missing: states.filter((state) => state.verdict === "missing").length,
			pass: states.filter((state) => state.verdict === "pass").length,
			review: states.filter((state) => state.verdict === "review").length,
		},
		timings,
	};
	await fs.writeFile(paths.report, JSON.stringify(report, null, "\t") + "\n");
	console.log(
		`${id}: ${JSON.stringify(report.summary)}; states ${String(slidev.length)} Slidev vs ${String(pptx.length)} pptx; ${String(clicks.problems.length)} problems.`,
	);
});

async function exists(file: string) {
	return fs.stat(file).then(
		() => true,
		() => false,
	);
}

async function pngs(directory: string) {
	return (await fs.readdir(directory))
		.filter((file) => file.endsWith(".png"))
		.sort()
		.map((file) => path.join(directory, file));
}

async function step(
	name: (typeof STEPS)[number],
	output: string,
	task: () => Promise<void>,
) {
	if (STEPS.indexOf(name) < from && (await exists(output))) {
		return;
	}

	const started = performance.now();
	await task();
	timings[name] = Math.round((performance.now() - started) / 100) / 10;
}
