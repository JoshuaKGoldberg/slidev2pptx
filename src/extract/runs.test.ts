import { describe, expect, it } from "vitest";

import { contentAreaOf } from "./runs.ts";

describe(contentAreaOf, () => {
	it.each([
		// Menlo at 20px: 18.56 and 4.72 round to 19 and 5.
		["Menlo", 18.56, 4.72, { ascent: 19, descent: 5, shift: 0.08 }],
		// Helvetica at 20px: 15.4 and 4.6 round to 15 and 5, and macOS adds 15% of 20 to the ascent.
		["Helvetica", 15.4, 4.6, { ascent: 18, descent: 5, shift: 1.1 }],
		["Times", 17.8, 4.34, { ascent: 21, descent: 4, shift: 1.77 }],
		// Only Courier, Helvetica and Times themselves get more ascent.
		["Helvetica Neue", 19.04, 4.6, { ascent: 19, descent: 5, shift: -0.22 }],
	])(
		"rounds %s's content area as Blink does",
		(font, ascent, descent, area) => {
			const { shift, ...rounded } = contentAreaOf(font, ascent, descent);

			expect(rounded).toEqual({ ascent: area.ascent, descent: area.descent });
			expect(shift).toBeCloseTo(area.shift, 2);
		},
	);
});
