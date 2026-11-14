import { execFile } from "node:child_process";
import { promisify } from "node:util";

const run = promisify(execFile);

/** Size both images are compared at: small enough to forgive anti-aliasing, big enough to see a misplaced word. */
const COMPARE_SIZE = "960x540!";

export interface ImageComparison {
	/** The fraction of pixels that differ by more than the fuzz, 0-1. */
	changed: number;
	/** Structural similarity, 0-1, where 1 is identical. */
	ssim: number;
}

/** Compares two renderings, writing a heatmap of their differences. */
export async function compareImages(
	expected: string,
	actual: string,
	heatmap: string,
): Promise<ImageComparison> {
	const ssim = 1 - (await metric("SSIM", expected, actual, "null:"));
	const changed = await metric("AE", expected, actual, heatmap, [
		"-fuzz",
		"10%",
	]);
	return { changed: changed / (960 * 540), ssim };
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
			"(",
			expected,
			"-resize",
			COMPARE_SIZE,
			")",
			"(",
			actual,
			"-resize",
			COMPARE_SIZE,
			")",
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
