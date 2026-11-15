import { parseHTML } from "linkedom";
import { describe, expect, it } from "vitest";

import type { ExtractContext, TextRuns } from "./context.ts";

import { baselineOf, lineShiftsOf } from "./runs.ts";

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

describe(lineShiftsOf, () => {
	// Metrics measured from Helvetica and Menlo.
	const METRICS: Record<string, { ascent: number; descent: number }> = {
		Helvetica: { ascent: 0.77, descent: 0.23 },
		Menlo: { ascent: 0.9282, descent: 0.2358 },
	};
	const box = {
		height: 0,
		matrix: [1, 0, 0, 1] as [number, number, number, number],
		width: 400,
		x: 0,
		y: 0,
	};

	/** Lays out each paragraph of a body as one line per run, 26px apart, as takumi does for Helvetica with Menlo at 20px. */
	const setUp = (body: string) => {
		const { document } = parseHTML(`<html><body>${body}</body></html>`);
		const elements = [...document.querySelectorAll("*")];
		const ids = new Map(
			elements.map((element, index) => [element, String(index)]),
		);
		const entries = new Map<Element, TextRuns>();
		let y = 0;
		for (const paragraph of document.querySelectorAll("p")) {
			const runs = [...paragraph.children].map((span) => {
				const run = {
					height: 20,
					owner: ids.get(span) ?? "",
					text: "x",
					width: 10,
					x: 0,
					y,
				};
				y += 26;
				return run;
			});
			entries.set(paragraph, { box, element: paragraph, runs });
		}

		const context = {
			byId: new Map(elements.map((element) => [ids.get(element), element])),
			contentInsets: () => ({ bottom: 0, left: 0, right: 0, top: 0 }),
			fonts: {
				metrics: (font: string) => METRICS[font],
				resolve: (family: string) => ({ font: family }),
			},
			looseRunsOf: () => [],
			ownRunsOf: (element: Element) => entries.get(element),
			styleOf: (element: Element) => {
				const style = element.getAttribute("style") ?? "";
				const inherited = (name: string): string | undefined => {
					for (
						let current: Element | null = element;
						current;
						current = current.parentElement
					) {
						const value = new RegExp(`(?:^|;)\\s*${name}:\\s*([^;]+)`).exec(
							current.getAttribute("style") ?? "",
						)?.[1];
						if (value) {
							return value.trim();
						}
					}
				};

				const own = (name: string) =>
					new RegExp(`(?:^|;)\\s*${name}:\\s*([^;]+)`).exec(style)?.[1]?.trim();
				return {
					display:
						own("display") ??
						(element.localName === "span" ? "inline" : "block"),
					"flex-direction": own("flex-direction") ?? "row",
					"font-family": inherited("font-family") ?? "Helvetica",
					"font-size": "20px",
					"font-style": "normal",
					"font-weight": "400",
					height: own("height"),
					"justify-content": own("justify-content") ?? "normal",
					"line-height": "24px",
					position: "static",
					"vertical-align": "baseline",
				};
			},
		} as unknown as ExtractContext;
		const shiftsOf = (paragraph: Element) => {
			const entry = entries.get(paragraph);
			if (!entry) {
				throw new Error("No such paragraph.");
			}

			const shiftOf = lineShiftsOf(context, entry, paragraph);
			return entry.runs.map((run) => shiftOf(run));
		};

		return { document, shiftsOf };
	};

	const HELVETICA = `<p><span>Helvetica</span></p>`;
	const MIXED = `<p><span style="font-family: Menlo">Menlo</span></p>`;

	it("drops a Helvetica line by Blink's taller ascent", () => {
		const { document, shiftsOf } = setUp(HELVETICA);

		expect(shiftsOf(document.querySelectorAll("p")[0])).toEqual([1]);
	});

	it("leaves a Menlo line in a Helvetica paragraph where takumi puts it", () => {
		const { document, shiftsOf } = setUp(MIXED);

		expect(shiftsOf(document.querySelectorAll("p")[0])).toEqual([0]);
	});

	it.each([
		// takumi's mixed line is 1px taller than Blink's: in a column that stacks from the start, nothing above it moves.
		["flex-start", [1, 0]],
		// From the end, the line above it comes down 1px.
		["flex-end", [2, 1]],
		// From the center, everything comes down half that.
		["center", [1.5, 0.5]],
	])(
		"moves lines stacked %s with a mixed line by its growth",
		(justify, shifts) => {
			const { document, shiftsOf } = setUp(
				`<div style="display: flex; flex-direction: column; height: 200px; justify-content: ${justify}">${HELVETICA}${MIXED}</div>`,
			);
			const [first, second] = document.querySelectorAll("p");

			expect([...shiftsOf(first), ...shiftsOf(second)]).toEqual(shifts);
		},
	);

	it("moves later lines of a paragraph up by an earlier mixed line's growth", () => {
		const { document, shiftsOf } = setUp(
			`<p><span style="font-family: Menlo">Menlo</span><span>Helvetica</span></p>`,
		);

		expect(shiftsOf(document.querySelectorAll("p")[0])).toEqual([0, 0]);
	});

	it("ignores justification in a flex column sized by its content", () => {
		const { document, shiftsOf } = setUp(
			`<div style="display: flex; flex-direction: column; justify-content: flex-end">${HELVETICA}${MIXED}</div>`,
		);
		const [first, second] = document.querySelectorAll("p");

		expect([...shiftsOf(first), ...shiftsOf(second)]).toEqual([1, 0]);
	});
});
