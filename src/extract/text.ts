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
	lineFontAreaOf,
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
	const { shift } = lineFontAreaOf(context, pre);
	const lines = [...pre.querySelectorAll(CODE_LINES)];

	const lineOf = (run: LaidOutRun, owner: Element) =>
		lines.length
			? lines.findLastIndex((line) => line.contains(owner))
			: Math.round(run.y / lineHeight);
	const placed = [pre, ...pre.querySelectorAll("*")]
		.flatMap((element) => {
			const own = context.ownRunsOf(element);
			return [...(own ? [own] : []), ...context.looseRunsOf(element)];
		})
		.flatMap((laidOutRuns) => {
			const frame = runFrame(context, laidOutRuns);
			const rotated = geometryOf(laidOutRuns.box).rotation !== 0;
			return laidOutRuns.runs.flatMap((run) => {
				const owner = context.byId.get(run.owner);
				const index = owner ? lineOf(run, owner) : -1;
				const clip = owner && !rotated && context.clipOf(owner, true);
				return owner &&
					index >= 0 &&
					!isLineFeed(run) &&
					(!clip || showsIn(runRect(frame, run), clip))
					? [{ frame, index, owner, run }]
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
	for (const { frame, index, owner, rect, run } of placed
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

/** Inline elements with a background or border ("pills", like inline code) as boxes, one per line they span. */
export function inlinePills(
	context: ExtractContext,
	element: Element,
	entry: TextRuns,
) {
	const { origin, scale } = runFrame(context, entry);
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
		const label = context.labelOf(child);
		for (const [index, line] of lines.entries()) {
			const first = index === 0;
			const last = index === lines.length - 1;
			const left = line.left - (first ? insets.left : 0);
			const right = line.right + (last ? insets.right : 0);
			const top = line.top - insets.top;
			context.push(
				child,
				...boxItems(style, {
					box: {
						height: (line.bottom + insets.bottom - top) * scale,
						width: (right - left) * scale,
						x: origin.x + left * scale,
						y: origin.y + top * scale,
					},
					flip: false,
					key: `${pathOf(child)}:${String(Math.round(line.left))}`,
					label,
					rotation: 0,
					scale,
					sides: SIDES.filter(
						(side) => (side !== "left" || first) && (side !== "right" || last),
					),
				}),
			);
		}
	}
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
	const { baselines, right, runs, start, wrapRight } = runsToParagraph(
		context,
		entry,
		element,
	);
	const paragraphs = [{ runs }];
	if (isInvisible(paragraphs)) {
		return;
	}

	const style = context.styleOf(element);
	const rtl = isRightToLeft(context, element);
	// CSS start-aligns a line too wide for its box, where PowerPoint would center or end-align it.
	const overflows =
		!rtl &&
		baselines.length === 1 &&
		Math.abs(start - content.x) <= OVERFLOW_TOLERANCE &&
		right - (content.x + content.width) > OVERFLOW_TOLERANCE;
	// Unturned, the box widens to the line, which PowerPoint would otherwise break; turned, it keeps the center it turns about.
	const box =
		overflows && !rotation ? { ...content, width: right - content.x } : content;
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
			wrapWidth: wrapRight === undefined ? undefined : wrapRight - box.x,
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
	const { shift } = lineFontAreaOf(context, entry.element ?? container);
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
				bottom: area.baseline + shift + area.descent,
				top: area.baseline + shift - area.ascent,
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
