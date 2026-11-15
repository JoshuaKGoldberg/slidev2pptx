import { execFile } from "node:child_process";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { promisify } from "node:util";

const run = promisify(execFile);

/**
 * Both images are shrunk and lightly blurred before comparing, which forgives
 * Chromium and PowerPoint anti-aliasing glyphs differently, but still catches
 * missing, moved, or restyled content.
 */
const NORMALIZE = ["-resize", "480x270!", "-blur", "0x1"];

const PIXELS = 480 * 270;

export interface ImageComparison {
	/** The fraction of pixels that differ by more than the fuzz, 0-1. */
	changed: number;
	/** Structural similarity, 0-1, where 1 is identical. */
	ssim: number;
}

/**
 * Thresholds on the normalized comparison, calibrated on states that look the
 * same (about 0.4% changed, 0.991 SSIM) and that visibly differ (over 1%).
 * They rank what reviewers look at; reviewers decide what's a real difference.
 */
export const THRESHOLDS = {
	changed: { fail: 0.03, review: 0.008 },
	ssim: { fail: 0.97, review: 0.985 },
};

export type Verdict = "fail" | "missing" | "pass" | "review";

export function verdictOf(comparison: ImageComparison | undefined): Verdict {
	return !comparison
		? "missing"
		: comparison.ssim < THRESHOLDS.ssim.fail ||
			  comparison.changed > THRESHOLDS.changed.fail
			? "fail"
			: comparison.ssim < THRESHOLDS.ssim.review ||
				  comparison.changed > THRESHOLDS.changed.review
				? "review"
				: "pass";
}

/** Compares two renderings, writing a heatmap of their differences. */
export async function compareImages(
	expected: string,
	actual: string,
	heatmap: string,
): Promise<ImageComparison> {
	const directory = await fs.mkdtemp(path.join(os.tmpdir(), "compare-"));
	try {
		const [a, b] = await Promise.all(
			[expected, actual].map(async (image, index) => {
				const normalized = path.join(directory, `${String(index)}.png`);
				await run("magick", [image, ...NORMALIZE, normalized]);
				return normalized;
			}),
		);
		const ssim = 1 - (await metric("SSIM", a, b, "null:"));
		const changed = await metric("AE", a, b, heatmap, ["-fuzz", "15%"]);
		return { changed: changed / PIXELS, ssim };
	} finally {
		await fs.rm(directory, { force: true, recursive: true });
	}
}

async function metric(
	name: string,
	expected: string,
	actual: string,
	output: string,
	extra: string[] = [],
) {
	try {
		const { stderr } = await run("magick", [
			"compare",
			"-metric",
			name,
			...extra,
			expected,
			actual,
			output,
		]);
		return parseMetric(name, stderr);
	} catch (error) {
		// compare exits with 1 when the images differ, which is fine.
		const { code, stderr } = error as { code?: number; stderr?: string };
		if (code === 1 && stderr) {
			return parseMetric(name, stderr);
		}

		throw error;
	}
}

function parseMetric(name: string, output: string) {
	const match = /^\s*([\d.e+-]+)(?: \(([\d.e+-]+)\))?/.exec(output);
	if (!match) {
		throw new Error(`Couldn't read the ${name} metric: ${output}`);
	}

	return Number(name === "AE" ? match[1] : match[2]);
}
