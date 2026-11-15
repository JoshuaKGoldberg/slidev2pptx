import fs from "node:fs/promises";
import path from "node:path";

import { renderSlidesWithPowerPoint } from "../render/powerpoint.ts";
import { compareImages, type ImageComparison, verdictOf } from "./images.ts";

/**
 * How many more times to render states that didn't pass. PowerPoint for Mac
 * now and then lays out a paragraph without its font's kerning, a few percent
 * wider, and which paragraphs it picks changes from one render of the same
 * file to the next.
 */
const RERENDERS = 2;

export interface RenderedState {
	/** The state's PowerPoint render, replaced when a later render is closer. */
	actual: string;
	comparison: ImageComparison;
	/** Where the state's heatmap goes. */
	diff: string;
	/** The state's Slidev render. */
	expected: string;
	/** The state's slide among the presentation's shown slides, from 0. */
	index: number;
	/** Which render was kept, from 1. */
	render: number;
}

/**
 * Renders states that didn't pass again, keeping each state's closest render:
 * a real difference shows up in every render, and PowerPoint's noise doesn't.
 * @param pptx The presentation the states were rendered from.
 * @param states The states to check.
 * @param scratch A directory to render into, removed afterwards.
 */
export async function rerenderMisses(
	pptx: string,
	states: RenderedState[],
	scratch: string,
) {
	try {
		for (let render = 2; render <= RERENDERS + 1; render++) {
			const misses = states.filter(
				(state) => verdictOf(state.comparison) !== "pass",
			);
			if (!misses.length) {
				return;
			}

			const directory = path.join(scratch, String(render));
			const files = await renderSlidesWithPowerPoint(
				pptx,
				misses.map((state) => state.index),
				directory,
			);
			for (const [position, state] of misses.entries()) {
				const file = files[position];
				if (!file) {
					continue;
				}

				const diff = path.join(directory, `${String(position)}-diff.png`);
				const comparison = await compareImages(state.expected, file, diff);
				if (closer(comparison, state.comparison)) {
					await fs.copyFile(file, state.actual);
					await fs.copyFile(diff, state.diff);
					Object.assign(state, { comparison, render });
				}
			}
		}
	} finally {
		await fs.rm(scratch, { force: true, recursive: true });
	}
}

const RANKS = { fail: 2, missing: 3, pass: 0, review: 1 };

function closer(next: ImageComparison, previous: ImageComparison) {
	const ranks = [next, previous].map(
		(comparison) => RANKS[verdictOf(comparison)],
	);
	return ranks[0] !== ranks[1]
		? ranks[0] < ranks[1]
		: next.ssim > previous.ssim;
}
