import { describe, expect, it } from "vitest";

import type { ExtractContext } from "./context.ts";

import { baselineOf } from "./runs.ts";

describe(baselineOf, () => {
	const context = {
		fonts: { metrics: () => ({ ascent: 0.75, descent: 0.25 }) },
	} as unknown as ExtractContext;
	const frame = { origin: { x: 0, y: 10 }, scale: 1 };
	const run = (text: string) => ({
		height: 40,
		owner: "1",
		text,
		width: 40,
		x: 0,
		y: 0,
	});

	it.each([
		// The text font's ascent share of its box.
		["Text", 10 + 0.75 * 40],
		// takumi's emoji box, whatever the text's font.
		["🤩", 10 + 0.8206 * 40],
	])("puts %j's baseline where takumi does", (text, baseline) => {
		expect(baselineOf(context, frame, run(text), "Body", 0)).toBeCloseTo(
			baseline,
		);
	});

	it("moves a baseline by Blink's shift, scaled into the slide", () => {
		expect(
			baselineOf(context, { ...frame, scale: 2 }, run("Text"), "Body", 1.5),
		).toBeCloseTo(10 + 0.75 * 80 + 3);
	});
});
