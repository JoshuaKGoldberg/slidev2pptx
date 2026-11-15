import type { LaidOutRun } from "../layout.ts";
import type { Paragraph } from "../model.ts";
import type { ExtractContext, TextRuns } from "./context.ts";

import { lineHeightOf } from "../cascade.ts";
import { availableSpanOf, boxItems, boxStyle } from "./boxes.ts";
import { pathOf } from "./context.ts";
import { geometryOf, insetBox, showsIn, SIDES } from "./geometry.ts";
import {
	alignOf,
	baselineOf,
	isInvisible,
	isLineFeed,
	isRightToLeft,
	lineShiftsOf,
	padRun,
	runAreaOf,
	runFrame,
	runRect,
	runsToParagraph,
	runStyle,
	signatureOf,
	startsLine,
} from "./runs.ts";

/** Link targets PowerPoint can open. */
const LINK_PROTOCOL = /^(?:https?|mailto):/i;

/** px a line may run past its box's end before it counts as overflowing, for layout rounding. */
const OVERFLOW_TOLERANCE = 0.5;

/** A code block's lines, in order: highlighted lines, twoslash error lines, and twoslash query popups, which show below their token's line. */
const CODE_LINES = ".line, .twoslash-meta-line, .twoslash-popup-container";

export function codeItem(context: ExtractContext, pre: Element) {
	const laidOut = context.boxOf(pre);
	if (!laidOut || !context.ownRunsOf(pre)) {
		return;
	}

	const { box, scale } = geometryOf(laidOut);
	const lineHeight = lineHeightOf(context.styleOf(pre));
	const lines = [...pre.querySelectorAll(CODE_LINES)];

	const lineOf = (run: LaidOutRun, owner: Element) =>
		lines.length
			? lines.findLastIndex((line) => line.contains(owner))
			: Math.round(run.y / lineHeight);
	const placed = [pre, ...pre.querySelectorAll("*")]
		.flatMap((element) => {
			const own = context.ownRunsOf(element);
			const entries: TextRuns[] = [
				...(own ? [own] : []),
				...context.looseRunsOf(element),
			];
			return entries.map((laidOutRuns) => ({ element, laidOutRuns }));
		})
		.flatMap(({ element, laidOutRuns }) => {
			const frame = runFrame(context, laidOutRuns);
			const rotated = geometryOf(laidOutRuns.box).rotation !== 0;
			const shiftOf = lineShiftsOf(
				context,
				laidOutRuns,
				laidOutRuns.element ?? element,
			);
			return laidOutRuns.runs.flatMap((run) => {
				const owner = context.byId.get(run.owner);
				const index = owner ? lineOf(run, owner) : -1;
				const clip = owner && !rotated && context.clipOf(owner, true);
				return owner &&
					index >= 0 &&
					!isLineFeed(run) &&
					(!clip || showsIn(runRect(frame, run), clip))
					? [{ frame, index, owner, run, shift: shiftOf(run) }]
					: [];
			});
		});
	const paragraphs: Paragraph[] = Array.from(
		{ length: Math.max(lines.length, ...placed.map(({ index }) => index + 1)) },
		() => ({ runs: [] }),
	);
	const content = insetBox(box, context.contentInsets(pre), scale);
	const baselines: (number | undefined)[] = [];
	const ends: number[] = [];
	for (const { frame, index, owner, rect, run, shift } of placed
		.map((entry) => ({ ...entry, rect: runRect(entry.frame, entry.run) }))
		.toSorted(
			(first, second) =>
				first.index - second.index ||
				(Math.abs(first.rect.y - second.rect.y) > first.rect.height / 2
					? first.rect.y - second.rect.y
					: first.rect.x - second.rect.x),
		)) {
		const style = runStyle(context, owner, context.opacityOf(owner));
		const pad = padRun(context, style, rect.x - (ends[index] ?? content.x));
		paragraphs[index].runs.push(...(pad ? [pad] : []), {
			...style,
			text: run.text.replace(/\t/g, "  "),
		});
		ends[index] = rect.x + rect.width;
		baselines[index] ??= baselineOf(context, frame, run, style.font, shift);
	}

	while (paragraphs.at(-1)?.runs.length === 0) {
		paragraphs.pop();
	}

	if (isInvisible(paragraphs)) {
		return;
	}

	const measured = baselines.filter((baseline) => baseline !== undefined);
	const leadingBlankLines = baselines.findIndex(
		(baseline) => baseline !== undefined,
	);
	context.push(pre, {
		baselines:
			leadingBlankLines > 0
				? [measured[0] - leadingBlankLines * lineHeight * scale, ...measured]
				: measured,
		box: content,
		key: `${pathOf(pre)}#code:${signatureOf(paragraphs)}`,
		kind: "code",
		paragraphs,
		style: {
			align: "left",
			available: availableSpanOf(context, pre),
			lineHeight: lineHeight * scale,
			nowrap: true,
			rtl: false,
		},
	});
}

/**
 * Inline elements with a background or border ("pills", like inline code) as
 * boxes, one per line they span. Abutting sibling runs painted alike with only
 * a plain fill, like highlighted tokens, share one box, as Chromium paints them
 * as one band: separate boxes would add seams and shapes.
 */
export function inlinePills(
	context: ExtractContext,
	element: Element,
	entry: TextRuns,
) {
	const { origin, scale } = runFrame(context, entry);
	const pills: Pill[] = [];
	let latest: Pill | undefined;
	for (const child of element.querySelectorAll("*")) {
		const style =
			context.styleOf(child).display === "inline"
				? boxStyle(context, child, context.opacityOf(child))
				: undefined;
		if (!style) {
			continue;
		}

		const lines = lineBoundsOf(context, entry, child, element);
		const insets = context.contentInsets(child);
		for (const [index, line] of lines.entries()) {
			const first = index === 0;
			const last = index === lines.length - 1;
			const pill = {
				bottom: line.bottom + insets.bottom,
				child,
				first,
				last,
				left: line.left - (first ? insets.left : 0),
				right: line.right + (last ? insets.right : 0),
				style,
				top: line.top - insets.top,
			};
			if (latest && abuts(latest, pill)) {
				latest.right = pill.right;
				latest.last = pill.last;
			} else {
				pills.push(pill);
				latest = pill;
			}
		}
	}

	for (const pill of pills) {
		context.push(
			pill.child,
			...boxItems(pill.style, {
				box: {
					height: (pill.bottom - pill.top) * scale,
					width: (pill.right - pill.left) * scale,
					x: origin.x + pill.left * scale,
					y: origin.y + pill.top * scale,
				},
				flip: false,
				key: `${pathOf(pill.child)}:${String(Math.round(pill.left))}`,
				label: context.labelOf(pill.child),
				rotation: 0,
				scale,
				sides: SIDES.filter(
					(side) =>
						(side !== "left" || pill.first) && (side !== "right" || pill.last),
				),
			}),
		);
	}
}

/** One line's box behind an inline element, in its text entry's run space. */
interface Pill {
	bottom: number;
	child: Element;
	first: boolean;
	last: boolean;
	left: number;
	right: number;
	style: NonNullable<ReturnType<typeof boxStyle>>;
	top: number;
}

/** How far apart, in CSS px, two pills' edges may be and still count as touching. */
const PILL_TOLERANCE = 0.5;

/** The solid color a pill is filled with, if that's all it paints (no border, rounding, shadow or gradient). */
function plainFillOf({ style }: Pill) {
	return !style.border &&
		!style.edges &&
		!style.radius &&
		!style.shadow &&
		style.fill &&
		"color" in style.fill
		? style.fill
		: undefined;
}

/** Whether a pill continues another on its line, both siblings filled alike and plainly. */
function abuts(previous: Pill, next: Pill) {
	const before = plainFillOf(previous);
	const after = plainFillOf(next);
	return (
		previous.child.parentElement === next.child.parentElement &&
		after !== undefined &&
		before?.color === after.color &&
		before.transparency === after.transparency &&
		Math.abs(previous.top - next.top) < PILL_TOLERANCE &&
		Math.abs(previous.bottom - next.bottom) < PILL_TOLERANCE &&
		Math.abs(previous.right - next.left) < PILL_TOLERANCE
	);
}

/**
 * Transparent hotspots over links that show no underline, one per line, which
 * carry the link instead of its text: PowerPoint underlines any linked text.
 */
export function linkHotspots(
	context: ExtractContext,
	element: Element,
	entry: TextRuns,
) {
	if (geometryOf(entry.box).rotation) {
		return;
	}

	const { origin, scale } = runFrame(context, entry);
	const outer = element.closest("a[href]");
	for (const link of new Set([
		...(outer ? [outer] : []),
		...element.querySelectorAll("a[href]"),
	])) {
		const href = link.getAttribute("href") ?? "";
		const owner = entry.runs
			.map((run) => context.byId.get(run.owner))
			.find((owner) => owner && link.contains(owner));
		if (
			!owner ||
			!LINK_PROTOCOL.test(href) ||
			runStyle(context, owner, 1).underline
		) {
			continue;
		}

		for (const line of lineBoundsOf(context, entry, link, element)) {
			context.push(link, {
				box: {
					height: (line.bottom - line.top) * scale,
					width: (line.right - line.left) * scale,
					x: origin.x + line.left * scale,
					y: origin.y + line.top * scale,
				},
				fill: { color: "FFFFFF", transparency: 100 },
				href,
				key: `${pathOf(link)}#link:${String(Math.round(line.left))}`,
				kind: "box",
				radius: 0,
				rotation: 0,
			});
		}
	}
}

export function textItem(
	context: ExtractContext,
	element: Element,
	entry: TextRuns,
) {
	const { box: entryBox, rotation, scale } = geometryOf(entry.box);
	const content = entry.element
		? insetBox(entryBox, context.contentInsets(entry.element), scale)
		: entryBox;
	const { baselines, left, right, runs, start, wrapLimit, wrapRight } =
		runsToParagraph(context, entry, element);
	const paragraphs = [{ runs }];
	if (isInvisible(paragraphs)) {
		return;
	}

	const style = context.styleOf(element);
	const rtl = isRightToLeft(context, element);
	// Runs lie along the (turned) line from the content box's start, `left`, so the line is measured from there.
	const lineWidth = right - left;
	// CSS start-aligns a line too wide for its box, where PowerPoint would center or end-align it.
	const overflows =
		!rtl &&
		baselines.length === 1 &&
		Math.abs(start - left) <= OVERFLOW_TOLERANCE &&
		lineWidth - content.width > OVERFLOW_TOLERANCE &&
		// takumi lays out a ZWJ emoji sequence like 😮‍💨 as each of its emoji, where browsers draw one.
		!runs.some((run) => /\p{Join_Control}/u.test(run.text));
	// Unturned, the box widens to the line, which PowerPoint would otherwise break; turned, it keeps the center it turns about.
	const box =
		overflows && !rotation ? { ...content, width: lineWidth } : content;
	context.push(element, {
		baselines: rotation ? undefined : baselines,
		box,
		heading: /^h[1-6]$/.test(element.localName),
		key: `${pathOf(element)}#${entry.element ? "text" : "loose"}:${signatureOf(paragraphs)}`,
		kind: "text",
		paragraphs,
		rotation,
		style: {
			align: overflows ? "left" : alignOf(context, element),
			available: availableSpanOf(context, element),
			lineHeight: lineHeightOf(style) * scale,
			nowrap:
				/^(?:nowrap|pre)$/.test(style["white-space"]) || Boolean(rotation),
			rtl,
			wrapLimit,
			wrapWidth: wrapRight === undefined ? undefined : wrapRight - left,
		},
	});
}

/**
 * The bounds of an element's laid-out runs in a text entry, one per line they
 * span, in the entry's run space: as wide as takumi lays them out, and as
 * tall as the content area Blink paints their inline backgrounds over.
 */
function lineBoundsOf(
	context: ExtractContext,
	entry: TextRuns,
	element: Element,
	container: Element,
) {
	const shiftOf = lineShiftsOf(context, entry, entry.element ?? container);
	const own = entry.runs.flatMap((run) => {
		const owner = context.byId.get(run.owner);
		if (!owner || !element.contains(owner)) {
			return [];
		}

		const area = runAreaOf(
			context.fonts,
			run,
			runStyle(context, owner, 1).font,
		);
		return [
			{
				...run,
				bottom: area.baseline + shiftOf(run) + area.descent,
				top: area.baseline + shiftOf(run) - area.ascent,
			},
		];
	});
	const lines: { bottom: number; left: number; right: number; top: number }[] =
		[];
	for (const [index, run] of own.entries()) {
		let line = lines.at(-1);
		if (!line || startsLine(own[index - 1], run)) {
			line = {
				bottom: -Infinity,
				left: Infinity,
				right: -Infinity,
				top: Infinity,
			};
			lines.push(line);
		}

		line.left = Math.min(line.left, run.x);
		line.right = Math.max(line.right, run.x + run.width);
		line.top = Math.min(line.top, run.top);
		line.bottom = Math.max(line.bottom, run.bottom);
	}

	return lines;
}
