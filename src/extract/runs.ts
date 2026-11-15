import type { ComputedStyle } from "../cascade.ts";
import type { Color } from "../color.ts";
import type { FontBook } from "../fonts.ts";
import type { LaidOutRun } from "../layout.ts";
import type { Paragraph, Point, Run, TextStyle, Underline } from "../model.ts";
import type { ExtractContext, TextRuns } from "./context.ts";

import { isAtomicInline, isOutOfFlow, lineHeightOf } from "../cascade.ts";
import {
	parseColor,
	parseLinearGradient,
	parseShadows,
	transparencyOf,
} from "../color.ts";
import { contentAreaOf, inlineExtentOf, isEmojiText } from "../fonts.ts";
import { ZERO_WIDTH_SPACE } from "../layout.ts";
import { clipsToText, paintedSizeOf } from "./boxes.ts";
import { geometryOf, scaleOf, showsIn } from "./geometry.ts";

/** Where a text entry's runs sit: takumi reports them in local px from the content box's corner. */
interface RunFrame {
	origin: Point;
	scale: number;
}

const TEXT_ALIGNMENTS: Partial<Record<string, TextStyle["align"]>> = {
	center: "center",
	justify: "justify",
	left: "left",
	right: "right",
};

/** White space that HTML collapses, which leaves out no-break spaces. */
const COLLAPSIBLE_SPACE = /[ \t\n\r\f]+/g;

/** Removes the white space HTML collapses from the start of text, which leaves no-break spaces. */
const trimCollapsibleStart = (text: string) =>
	text.replace(/^[ \t\n\r\f]+/, "");

/** Removes the white space HTML collapses from the end of text, which leaves no-break spaces. */
const trimCollapsibleEnd = (text: string) => text.replace(/[ \t\n\r\f]+$/, "");

/** Turns trailing spaces into no-break spaces, which PowerPoint counts in a line's alignment. */
const holdTrailingSpaces = (text: string) =>
	text.replace(/ +$/, (spaces) => NO_BREAK_SPACE.repeat(spaces.length));

/** Gaps between runs narrower than this, in px, are kerning rather than spacing. */
const MIN_SPACER_GAP = 0.75;

/** Pair kerning smaller than this, in px, between runs is left out. */
const MIN_KERNING = 0.25;

export const NO_BREAK_SPACE = "\u00A0";

const PRESERVED_NEWLINES = /^(?:pre|pre-wrap|pre-line|break-spaces)$/;

const PRESERVED_SPACES = /^(?:pre|pre-wrap|break-spaces)$/;

const UNDERLINE_STYLES: Partial<Record<string, Underline["style"]>> = {
	dashed: "dashed",
	dotted: "dotted",
	double: "double",
	solid: "single",
	wavy: "wavy",
};

export function alignOf(context: ExtractContext, element: Element) {
	const rtl = isRightToLeft(context, element);
	const logical: Partial<Record<string, string>> = {
		end: rtl ? "left" : "right",
		start: rtl ? "right" : "left",
	};
	const align = context.styleOf(element)["text-align"];
	return TEXT_ALIGNMENTS[logical[align] ?? align] ?? "left";
}

/** Whether a CSS paint-order paints the stroke before the fill: the first of them it names, as it paints those it leaves out after in their usual order. */
const strokesBehind = (order: string) =>
	/\b(?:fill|stroke)\b/.exec(order)?.[0] === "stroke";

/** How far down its box takumi puts an emoji run's baseline, whatever the text's font (measured: 0.966em down a 1.178em box). */
const EMOJI_RUN_BASELINE = 0.8206;

const ascentRatio = (fonts: FontBook, font: string) => {
	const { ascent, descent } = fonts.metrics(font);
	return ascent / (ascent + descent);
};

/** How far above and below its baseline an element's inline box reaches in its lines, in layout px, as Blink and takumi lay it out. */
interface InlineExtents {
	blink: { above: number; below: number };
	takumi: { above: number; below: number };
}

const extents = new WeakMap<ComputedStyle, InlineExtents>();

function inlineExtentsOf(
	context: ExtractContext,
	element: Element,
): InlineExtents {
	const style = context.styleOf(element);
	const known = extents.get(style);
	if (known) {
		return known;
	}

	const lineHeight = lineHeightOf(style);
	const size = parseFloat(style["font-size"]);
	const { font } = context.fonts.resolve(
		style["font-family"],
		Number(style["font-weight"]),
		style["font-style"] === "italic",
	);
	const { ascent, descent } = context.fonts.metrics(font);
	const extentIn = (engine: "blink" | "takumi") =>
		inlineExtentOf(font, ascent * size, descent * size, lineHeight, engine);
	const extent = { blink: extentIn("blink"), takumi: extentIn("takumi") };
	extents.set(style, extent);
	return extent;
}

/** One line of a text entry, as Blink lays it out against takumi, in slide px. */
interface LineBox {
	/** How far below takumi's Blink puts the line's baseline, measured from the line's top. */
	drop: number;
	/** How much taller Blink's line is than takumi's. */
	growth: number;
	/** Where takumi puts the line's first run, which orders lines. */
	y: number;
}

/**
 * A text entry's lines as Blink lays them out against takumi. Each line
 * reaches as far above its baseline as the most any of its inline boxes (its
 * block's strut, and each run's element) does, and as far below. The two
 * engines lay each box out alike but for Blink's taller Helvetica and its kin
 * on macOS, so a line in one of those draws its baseline lower, and a line
 * mixing one with another font is shorter in Blink than in takumi.
 */
function lineBoxesOf(
	context: ExtractContext,
	entry: TextRuns,
	source: Element,
) {
	const frame = runFrame(context, entry);
	const strut = inlineExtentsOf(context, source);
	const groups: { extents: InlineExtents[]; y: number }[] = [];
	const lineOfRun = new Map<LaidOutRun, number>();
	let previous: LaidOutRun | undefined;
	for (const run of entry.runs) {
		const owner = context.byId.get(run.owner);
		if (!owner || isLineFeed(run)) {
			continue;
		}

		if (!groups.length || startsLine(previous, run)) {
			groups.push({ extents: [strut], y: runRect(frame, run).y });
		}

		previous = run;
		// Raised and lowered text sits apart from the baseline the line's other boxes share.
		if (owner !== source && !scriptOf(context.styleOf(owner), owner)) {
			groups[groups.length - 1].extents.push(inlineExtentsOf(context, owner));
		}

		lineOfRun.set(run, groups.length - 1);
	}

	const lines = groups.map(({ extents, y }): LineBox => {
		const most = (engine: keyof InlineExtents, side: "above" | "below") =>
			Math.max(...extents.map((extent) => extent[engine][side]));
		const drop = most("blink", "above") - most("takumi", "above");
		return {
			drop: drop * frame.scale,
			growth:
				(drop + most("blink", "below") - most("takumi", "below")) * frame.scale,
			y,
		};
	});
	return { lineOfRun, lines };
}

/** Lines this close, in slide px, are one line. */
const SAME_LINE = 1;

/**
 * The flex column a block's lines are stacked in, and where it puts the space
 * they leave: Blink's growth of the lines before a line moves it down from the
 * start, of the lines after it moves it up from the end, and half of all of
 * them moves it up from the center. Without a sized flex column that stacks
 * the block alone, the block's own lines are all that move it.
 */
function anchorOf(context: ExtractContext, source: Element) {
	for (
		let current = source.parentElement;
		current && current.id !== "slide-content";
		current = current.parentElement
	) {
		const style = context.styleOf(current);
		const display = style.display;
		if (/^(?:inline-)?flex$/.test(display)) {
			if (style["flex-direction"] !== "column") {
				return undefined;
			}

			const sized = !/^(?:auto)?$/.test(style.height ?? "");
			const justify = style["justify-content"] ?? "";
			return {
				container: current,
				justify: !sized
					? "start"
					: /^(?:flex-)?end$/.test(justify)
						? "end"
						: justify === "center"
							? "center"
							: "start",
			} as const;
		}

		if (/^(?:inline-)?grid$/.test(display) || isOutOfFlow(style)) {
			return undefined;
		}
	}

	return undefined;
}

const containerLines = new WeakMap<ExtractContext, Map<Element, LineBox[]>>();

/** Every line of text in a flex column, from each block in it, in any order. */
function containerLinesOf(context: ExtractContext, container: Element) {
	let cache = containerLines.get(context);
	if (!cache) {
		cache = new Map();
		containerLines.set(context, cache);
	}

	let lines = cache.get(container);
	if (!lines) {
		lines = [container, ...container.querySelectorAll("*")].flatMap(
			(element) => {
				const own = context.ownRunsOf(element);
				const entries: TextRuns[] = [
					...(own ? [own] : []),
					...context.looseRunsOf(element),
				];
				return entries
					.filter((entry) => !geometryOf(entry.box).rotation)
					.flatMap(
						(entry) =>
							lineBoxesOf(context, entry, entry.element ?? element).lines,
					);
			},
		);
		cache.set(container, lines);
	}

	return lines;
}

/**
 * How far below takumi's baseline Blink draws each run's, in layout px: its
 * line's drop, and how far the growth of the lines stacked with it moves it.
 */
export function lineShiftsOf(
	context: ExtractContext,
	entry: TextRuns,
	source: Element,
) {
	let bySource = lineShifts.get(entry.runs);
	if (!bySource) {
		bySource = new Map();
		lineShifts.set(entry.runs, bySource);
	}

	let shiftOf = bySource.get(source);
	if (!shiftOf) {
		shiftOf = measureLineShifts(context, entry, source);
		bySource.set(source, shiftOf);
	}

	return shiftOf;
}

const lineShifts = new WeakMap<
	LaidOutRun[],
	Map<Element, (run: LaidOutRun) => number>
>();

function measureLineShifts(
	context: ExtractContext,
	entry: TextRuns,
	source: Element,
) {
	const { lineOfRun, lines } = lineBoxesOf(context, entry, source);
	const { scale } = runFrame(context, entry);
	const anchor = geometryOf(entry.box).rotation
		? undefined
		: anchorOf(context, source);
	const stacked = anchor ? containerLinesOf(context, anchor.container) : lines;
	const growthOf = (test: (line: LineBox) => boolean) =>
		stacked.filter(test).reduce((sum, line) => sum + line.growth, 0);
	const shifts = lines.map((line) => {
		const before = growthOf((other) => other.y < line.y - SAME_LINE);
		const moved =
			anchor?.justify === "end"
				? -line.growth - growthOf((other) => other.y > line.y + SAME_LINE)
				: anchor?.justify === "center"
					? before - growthOf(() => true) / 2
					: before;
		return (line.drop + moved) / scale;
	});
	return (run: LaidOutRun) => shifts[lineOfRun.get(run) ?? -1] ?? 0;
}

/** The content area of a run's font, sized by the height takumi gave the run. */
export function runAreaOf(fonts: FontBook, run: LaidOutRun, font: string) {
	const ascent = run.height * ascentRatio(fonts, font);
	return {
		...contentAreaOf(font, ascent, run.height - ascent),
		/** Where takumi puts the run's baseline, in the entry's run space. */
		baseline: run.y + ascent,
	};
}

/** Where Blink draws a run's baseline, given how far below takumi's its line's font puts it in layout px. */
export const baselineOf = (
	context: ExtractContext,
	frame: RunFrame,
	run: LaidOutRun,
	font: string,
	shift: number,
) => {
	const rect = runRect(frame, run);
	// takumi lays emoji out in their own font, whose box sits differently on the baseline.
	const ratio = isEmojiText(run.text)
		? EMOJI_RUN_BASELINE
		: ascentRatio(context.fonts, font);
	return rect.y + rect.height * ratio + shift * frame.scale;
};

export const isLineFeed = (run: LaidOutRun) => /^\n+$/.test(run.text);

export const isRightToLeft = (context: ExtractContext, element: Element) =>
	context.styleOf(element).direction === "rtl";

export const isInvisible = (paragraphs: Paragraph[]) =>
	paragraphs.every(({ runs }) => runs.every((run) => run.transparency >= 100));

export const signatureOf = (paragraphs: Paragraph[]) =>
	paragraphs
		.map(({ runs }) =>
			runs
				.map(
					(run) =>
						`${run.text}|${run.color}|${String(run.transparency)}|${run.size.toFixed(1)}`,
				)
				.join(""),
		)
		.join("¶");

export function runFrame(context: ExtractContext, entry: TextRuns) {
	const scale = scaleOf(entry.box);
	const { left, top } = entry.element
		? context.contentInsets(entry.element)
		: { left: 0, top: 0 };
	return {
		origin: { x: entry.box.x + scale * left, y: entry.box.y + scale * top },
		scale,
	};
}

export const runRect = ({ origin, scale }: RunFrame, run: LaidOutRun) => ({
	height: run.height * scale,
	width: run.width * scale,
	x: origin.x + run.x * scale,
	y: origin.y + run.y * scale,
});

/** Converts takumi's runs (per line, per style) into one paragraph's runs, with baselines and soft breaks for browser wraps. */
export function runsToParagraph(
	context: ExtractContext,
	entry: TextRuns,
	container: Element,
	opacityOf = context.opacityOf,
) {
	const frame = runFrame(context, entry);
	const source = entry.element ?? container;
	const shiftOf = lineShiftsOf(context, entry, source);
	const clip = geometryOf(entry.box).rotation
		? undefined
		: context.clipOf(source, true);
	const runs: Run[] = [];
	const baselines: number[] = [];
	const matchSource = sourceMatcher(sourceTextOf(context, source));
	const padsLineStarts = /^(?:justify|left)$/.test(alignOf(context, source));
	let previous: LaidOutRun | undefined;
	let right = -Infinity;
	let start = Infinity;
	let line = { left: Infinity, right: -Infinity };
	/** Where the browser wrapped lines: how wide each line it ended was, and which run starts the next. */
	const wraps: { ended: number; next: number }[] = [];
	let wrapped = false;
	let afterNewline = false;
	let leadingSpaced = false;
	let trailingSpaced = false;
	// PowerPoint leaves trailing spaces out of a line's alignment, where browsers count preserved ones.
	const holdsTrailingSpaces = (text: string) =>
		trailingSpaced && !padsLineStarts
			? holdTrailingSpaces(text)
			: trimCollapsibleEnd(text);
	for (const [index, run] of entry.runs.entries()) {
		const owner = context.byId.get(run.owner);
		if (!owner || isLineFeed(run)) {
			afterNewline ||= run.text.includes("\n");
			continue;
		}

		const preserved = PRESERVED_SPACES.test(
			context.styleOf(owner)["white-space"],
		);
		const match = matchSource(run.text, preserved);
		const rect = runRect(frame, run);
		if (clip && !showsIn(rect, clip)) {
			continue;
		}

		const style = runStyle(context, owner, opacityOf(owner));
		right = Math.max(right, rect.x + rect.width);
		start = Math.min(start, rect.x);
		const newLine = startsLine(previous, run);
		const ended = line;
		line = newLine
			? { left: rect.x, right: rect.x + rect.width }
			: {
					left: Math.min(line.left, rect.x),
					right: Math.max(line.right, rect.x + rect.width),
				};
		const offset =
			(previous && !newLine
				? run.x - (previous.x + previous.width)
				: padsLineStarts
					? run.x
					: 0) * frame.scale;
		const pad = padRun(context, style, offset);
		previous = run;
		if (!style.script) {
			const baseline = baselineOf(
				context,
				frame,
				run,
				style.font,
				shiftOf(run),
			);
			if (!baselines.length || newLine) {
				baselines.push(baseline);
			} else {
				baselines[baselines.length - 1] = Math.max(
					baselines[baselines.length - 1],
					baseline,
				);
			}
		}

		const text = sourceCased(
			preserved ? run.text : run.text.replace(COLLAPSIBLE_SPACE, " "),
			style,
			match?.original,
		);
		if (!runs.length) {
			leadingSpaced = preserved;
		}

		if (newLine) {
			const last = runs.at(-1);
			const { blankLines, forced, spaced } = breakBefore(
				last?.text ?? "",
				match?.between,
				afterNewline,
			);
			wrapped ||= !forced;
			if (!forced) {
				wraps.push({ ended: ended.right - ended.left, next: index });
			}

			if (last && !last.pad) {
				// A space the browser wrapped at can be its own run, which the source match already passed.
				const hung = spaced || /[ \t\n\r\f]$/.test(last.text);
				last.text = forced
					? holdsTrailingSpaces(last.text)
					: trimCollapsibleEnd(last.text) + (hung ? " " : "");
			}

			for (let line = 0; line < blankLines; line++) {
				runs.push({
					...style,
					href: undefined,
					lineBreak: "forced",
					strike: false,
					text: " ",
					underline: undefined,
				});
			}

			// Preserved spaces indent a line the source starts, but hang at the end of a line the browser wraps.
			const indents = preserved && (forced || !last);
			const lineBreak = forced ? "forced" : "wrap";
			if (pad) {
				runs.push({ ...pad, lineBreak }, { ...style, text });
			} else {
				runs.push({
					...style,
					lineBreak,
					text: indents ? text : trimCollapsibleStart(text),
				});
			}
		} else {
			if (pad) {
				runs.push(pad);
			}

			if (!pad) {
				kernAcross(context, runs, style, text);
			}

			runs.push({ ...style, text });
		}

		trailingSpaced = preserved;
		afterNewline = false;
	}

	return {
		baselines,
		left: frame.origin.x,
		right,
		runs: trimRuns(runs, leadingSpaced, holdsTrailingSpaces),
		/** Where the text's first run starts. */
		start,
		wrapLimit:
			wrapped && wraps.length
				? Math.min(
						...wraps.map(
							({ ended, next }) =>
								ended + firstWordWidth(context, frame, entry.runs, next),
						),
					)
				: undefined,
		wrapRight: wrapped ? right : undefined,
	};
}

/**
 * How wide the first word of the line a run starts draws: up to the first
 * space it may break at (not a no-break space), across runs if it spans
 * several. Within a run, its share is by advance in the run's font where
 * that's measurable, else by characters.
 */
function firstWordWidth(
	context: ExtractContext,
	frame: RunFrame,
	runs: LaidOutRun[],
	start: number,
) {
	let width = 0;
	for (let index = start; index < runs.length; index++) {
		const run = runs[index];
		const owner = context.byId.get(run.owner);
		if (
			!owner ||
			isLineFeed(run) ||
			(index > start && startsLine(runs[index - 1], run))
		) {
			break;
		}

		const { font } = runStyle(context, owner, 1);
		const rect = runRect(frame, run);
		const text = index === start ? trimCollapsibleStart(run.text) : run.text;
		const word = /^[^ \t\n\r\f]*/.exec(text)?.[0] ?? "";
		const whole = context.fonts.advance(font, run.text);
		const part = context.fonts.advance(font, word);
		width +=
			rect.width *
			(whole && part !== undefined
				? part / whole
				: word.length / Math.max(1, run.text.length));
		if (word.length < text.length) {
			break;
		}
	}

	return width;
}

/** A spacer run as wide as a gap between runs, in a run's style; undefined when the gap is only kerning. */

/**
 * Keeps the pair kerning browsers apply across abutting runs in one font,
 * like a period colored apart from the quote before it, which PowerPoint
 * (and takumi) leave out between runs: the previous run's last character
 * gets the pair's kerning as letter spacing.
 */
export function padRun(
	context: ExtractContext,
	style: Omit<Run, "text">,
	offset: number,
): Run | undefined {
	const spacer = context.fonts.spacer();
	return offset > MIN_SPACER_GAP
		? {
				...style,
				bold: false,
				font: spacer.font,
				gradient: undefined,
				letterSpacing: offset - spacer.advance * style.size,
				pad: true,
				stroke: undefined,
				text: NO_BREAK_SPACE,
				transparency: 100,
			}
		: undefined;
}

export function runStyle(
	context: ExtractContext,
	element: Element,
	opacity: number,
): Omit<Run, "text"> {
	const style = context.styleOf(element);
	const zoom = context.scaleAt(element);
	const color = textColorOf(context, element);
	const gradient =
		!color || color.alpha === 0 ? gradientTextOf(context, element) : undefined;
	const visible = gradient ? { alpha: 1, hex: gradient.stops[0].color } : color;
	const { bold, font } = context.fonts.resolve(
		style["font-family"],
		Number(style["font-weight"]),
		style["font-style"] === "italic",
	);
	const stroke = context.lengthOf(element, "-webkit-text-stroke-width");
	const href = element.closest("a[href]")?.getAttribute("href") ?? "";
	const shadow = parseShadows(
		style["text-shadow"],
		(length) => context.resolve(element, length),
		style.color,
	).at(0)?.shadow;
	const decorating = decoratingStylesOf(context, element);
	return {
		bold,
		color: visible?.hex ?? "FFFFFF",
		font,
		gradient: gradient && {
			angle: gradient.angle,
			stops: gradient.stops.map((stop) => ({ ...stop, transparency: 0 })),
		},
		href: /^(?:https?|mailto):/i.test(href) ? href : undefined,
		italic: style["font-style"] === "italic",
		letterSpacing: context.lengthOf(element, "letter-spacing") * zoom,
		pad: false,
		script: scriptOf(style, element),
		shadow,
		size: parseFloat(style["font-size"]) * zoom,
		strike: decorating.some(({ "text-decoration-line": lines }) =>
			lines.includes("line-through"),
		),
		stroke: stroke
			? {
					behind: strokesBehind(style["paint-order"]),
					color:
						parseColor(style["-webkit-text-stroke-color"], style.color)?.hex ??
						"000000",
					width: stroke * zoom,
				}
			: undefined,
		transparency: visible ? transparencyOf(visible.alpha, opacity) : 100,
		underline: underlineOf(
			decorating.find(({ "text-decoration-line": lines }) =>
				lines.includes("underline"),
			),
			visible?.hex,
		),
		upper: style["text-transform"] === "uppercase",
	};
}

function kernAcross(
	context: ExtractContext,
	runs: Run[],
	next: Omit<Run, "text">,
	text: string,
) {
	const previous = runs.at(-1);
	if (
		!previous ||
		previous.pad ||
		previous.font !== next.font ||
		previous.size !== next.size ||
		previous.bold !== next.bold ||
		previous.italic !== next.italic ||
		previous.letterSpacing ||
		next.letterSpacing
	) {
		return;
	}

	const characters = Array.from(previous.text);
	const last = characters.pop();
	const first = Array.from(text).at(0);
	if (last === undefined || first === undefined || /\s/.test(last + first)) {
		return;
	}

	const kerning =
		(context.fonts.kerning(next.font, last, first) ?? 0) * next.size;
	if (Math.abs(kerning) < MIN_KERNING) {
		return;
	}

	const kerned = { ...previous, letterSpacing: kerning };
	if (characters.length) {
		previous.text = characters.join("");
		runs.push({ ...kerned, lineBreak: undefined, text: last });
	} else {
		runs[runs.length - 1] = kerned;
	}
}

/** Whether a run starts a new line: it jumps back left and down (runs in other fonts, like inline code, sit at other heights on one line). */
export const startsLine = (previous: LaidOutRun | undefined, run: LaidOutRun) =>
	previous !== undefined &&
	run.x < previous.x + previous.width - 1 &&
	run.y + run.height > previous.y + previous.height + 1;

/**
 * How a run that starts a new line breaks from the text before it. A br element or a
 * newline in preserved white space forces a break, as does a wrap at a zero-width space, which
 * PowerPoint can't break at; any other wrap keeps the space it wrapped at, if any.
 */
function breakBefore(
	previousText: string,
	between: string | undefined,
	afterNewline: boolean,
) {
	const newlines = between ? between.split("\n").length - 1 : 0;
	const breaks = Math.max(afterNewline ? 1 : 0, newlines);
	const trimmed = previousText.trimEnd();
	return {
		blankLines: Math.max(0, breaks - 1),
		forced: breaks > 0 || previousText.endsWith(ZERO_WIDTH_SPACE),
		spaced:
			between === undefined
				? !trimmed.endsWith("-") && !trimmed.endsWith(ZERO_WIDTH_SPACE)
				: /\s/.test(between),
	};
}

function decoratingStylesOf(context: ExtractContext, element: Element) {
	const styles: ComputedStyle[] = [];
	for (
		let current: Element | null = element;
		current && current.localName !== "html";
		current = current.parentElement
	) {
		const style = context.styleOf(current);
		styles.push(style);
		if (isAtomicInline(style) || isOutOfFlow(style)) {
			break;
		}
	}

	return styles;
}

function gradientTextOf(context: ExtractContext, element: Element) {
	for (
		let current: Element | null = element;
		current && current.localName !== "html";
		current = current.parentElement
	) {
		const style = context.styleOf(current);
		if (clipsToText(style)) {
			const gradient = parseLinearGradient(
				style["background-image"],
				1,
				paintedSizeOf(context, current),
				style.color,
			);
			if (gradient) {
				return gradient;
			}
		}
	}

	return undefined;
}

/** Raised or lowered text: by vertical-align, or inside sup or sub, which the Tailwind reset positions relatively instead. */
function scriptOf(style: ComputedStyle, element: Element) {
	if (style["vertical-align"] === "super" || element.closest("sup")) {
		return "super";
	}

	if (style["vertical-align"] === "sub" || element.closest("sub")) {
		return "sub";
	}

	return undefined;
}

/** Uppercase text in its source casing (takumi reports it transformed), which PowerPoint capitalizes. */
function sourceCased(
	text: string,
	style: Pick<Run, "upper">,
	original: string | undefined,
) {
	const core = text.trim();
	return style.upper && core.length === original?.length
		? text.replace(core, () => original)
		: text;
}

/**
 * Finds each run's text in turn in the source text, which keeps what takumi drops: the
 * casing before text-transform, and the white space or newlines lines broke at.
 */
function sourceMatcher(sourceText: string) {
	const source = sourceText.toLowerCase();
	let cursor = 0;
	return (text: string, preserved: boolean) => {
		const spaced = (
			preserved ? text : text.replace(COLLAPSIBLE_SPACE, " ")
		).toLowerCase();
		// A run of only preserved spaces, as takumi lays out an indent, matches as itself.
		const needle = spaced.trim() || (preserved ? spaced : "");
		const found = needle ? source.indexOf(needle, cursor) : -1;
		if (found < 0) {
			return undefined;
		}

		const between = source.slice(cursor, found);
		cursor = found + needle.length;
		return { between, original: sourceText.slice(found, cursor) };
	};
}

/** An element's text as browsers lay it out: collapsible white space as single spaces, and preserved newlines and br elements as "\n". */
function sourceTextOf(context: ExtractContext, root: Element | undefined) {
	let text = "";
	const visit = (node: Node) => {
		if (node.nodeType === node.TEXT_NODE) {
			const whiteSpace = node.parentElement
				? context.styleOf(node.parentElement)["white-space"]
				: "normal";
			const content = node.textContent ?? "";
			text += PRESERVED_SPACES.test(whiteSpace)
				? content
				: content.replace(
						PRESERVED_NEWLINES.test(whiteSpace) ? /[ \t]+/g : COLLAPSIBLE_SPACE,
						" ",
					);
		} else if (node.nodeName === "BR") {
			text += "\n";
		} else {
			node.childNodes.forEach(visit);
		}
	};

	if (root) {
		visit(root);
	}

	return text;
}

/** An element's text color; an invalid one falls back to its parent's, as browsers drop invalid declarations. */
function textColorOf(
	context: ExtractContext,
	element: Element,
): Color | undefined {
	const style = context.styleOf(element);
	const fill = style["-webkit-text-fill-color"];
	const value = fill && fill !== "currentcolor" ? fill : style.color;
	return (
		parseColor(value) ??
		(element.parentElement
			? textColorOf(context, element.parentElement)
			: undefined)
	);
}

function trimRuns(
	runs: Run[],
	leadingSpaced: boolean,
	trimEnd: (text: string) => string,
) {
	const first = runs.at(0);
	const last = runs.at(-1);
	if (first) {
		delete first.lineBreak;
		if (!first.pad && !leadingSpaced) {
			first.text = trimCollapsibleStart(first.text);
		}
	}

	if (last && !last.pad) {
		last.text = trimEnd(last.text);
	}

	return runs
		.map((run) => ({ ...run, text: run.text.replaceAll(ZERO_WIDTH_SPACE, "") }))
		.filter((run) => run.text.length);
}

function underlineOf(
	style: ComputedStyle | undefined,
	textColor: string | undefined,
) {
	if (!style) {
		return undefined;
	}

	const color = parseColor(style["text-decoration-color"], style.color)?.hex;
	return {
		color: color === textColor ? undefined : color,
		style: UNDERLINE_STYLES[style["text-decoration-style"]] ?? "single",
	};
}
