import type { ListParagraph, MarkerFace } from "../model.ts";
import type { ExtractContext, TextRuns } from "./context.ts";

import { hasMarker, lineHeightOf } from "../cascade.ts";
import { parseColor } from "../color.ts";
import { availableSpanOf, paintItems } from "./boxes.ts";
import { MIN_OPACITY, pathOf } from "./context.ts";
import { geometryOf, insetBox } from "./geometry.ts";
import { rasterItem } from "./images.ts";
import {
	alignOf,
	isInvisible,
	isRightToLeft,
	NO_BREAK_SPACE,
	runsToParagraph,
	runStyle,
} from "./runs.ts";
import { inlinePills, linkHotspots } from "./text.ts";

/** A piece of a list item: some of its text, or a nested list. */
type ItemBlock =
	| { entry: TextRuns; kind: "text"; y: number }
	| { kind: "list"; list: Element; y: number };

/** Whether an element is a list that exports as list paragraphs: a ul or ol whose grid or flex layout, if any, doesn't put items side by side, as cards do. */
export function isTextList(context: ExtractContext, element: Element) {
	if (element.localName !== "ul" && element.localName !== "ol") {
		return false;
	}

	if (!/\b(?:flex|grid)\b/.test(context.styleOf(element).display)) {
		return true;
	}

	const boxes = itemsOf(element).flatMap((item) => {
		const laidOut = context.boxOf(item);
		return laidOut ? [geometryOf(laidOut).box] : [];
	});
	return !boxes.some(
		(box, index) =>
			index > 0 && box.x >= boxes[index - 1].x + boxes[index - 1].width - 1,
	);
}

/** Whether an element is a list item that shows a marker but lays out apart from its list's text, as in a grid. */
export function isMarkedApart(context: ExtractContext, element: Element) {
	const list = element.parentElement;
	return (
		element.localName === "li" &&
		hasMarker(context.styleOf(element)) &&
		(list?.localName === "ol" || list?.localName === "ul") &&
		!isTextList(context, list)
	);
}

/** Whether a list item's child draws apart from its list's text: a block with no text of its own, like a code block, an image's paragraph, or a component. */
export const isListAside = (context: ExtractContext, child: Element) =>
	child.parentElement?.localName === "li" &&
	!isTextList(context, child) &&
	!context.isInline(child) &&
	context.ownRunsOf(child) === undefined;

/** Exports a list as list paragraphs, over its items' paint; returns the items it rendered as images instead, whose content it left out. */
export function listItem(
	context: ExtractContext,
	list: Element,
	opacity: number,
) {
	const rasterized: Element[] = [];
	const laidOut = context.boxOf(list);
	if (!laidOut) {
		return rasterized;
	}

	for (const item of itemsWithin(context, list)) {
		const itemBox = context.boxOf(item);
		const itemOpacity = context.opacityOf(item);
		if (
			!itemBox ||
			itemOpacity < MIN_OPACITY ||
			rasterized.some((raster) => raster.contains(item))
		) {
			continue;
		}

		const paint = paintItems(context, item, itemOpacity, geometryOf(itemBox));
		if (paint) {
			context.push(item, ...paint);
		} else {
			rasterItem(context, item);
			rasterized.push(item);
		}
	}

	const { box: listBox, scale } = geometryOf(laidOut);
	const content = insetBox(listBox, context.contentInsets(list), scale);
	const first = [...list.children].find((child) => child.localName === "li");
	const styled = first ?? list;
	const columns = columnsOf(context, list);
	for (const [index, { items, left, right }] of columns.entries()) {
		const x = content.x + left - columns[0].left;
		const box =
			columns.length > 1 ? { ...content, width: right - x, x } : content;
		const links: { element: Element; entry: TextRuns }[] = [];
		const converted = listParagraphs(
			context,
			list,
			0,
			box.x,
			opacity,
			rasterized,
			links,
			items,
		);
		const paragraphs = converted.map(({ paragraph }) => paragraph);
		if (isInvisible(paragraphs)) {
			continue;
		}

		context.push(list, {
			box,
			key: `${pathOf(list)}#list${index ? `-${String(index)}` : ""}`,
			kind: "list",
			paragraphs,
			style: {
				align: alignOf(context, styled),
				available:
					columns.length > 1
						? { left: x, right }
						: availableSpanOf(context, list),
				lineHeight: lineHeightOf(context.styleOf(styled)) * scale,
				nowrap: false,
				rtl: isRightToLeft(context, styled),
				wrapLimit: limitOf(converted.map(({ wrapLimit }) => wrapLimit)),
				wrapWidth: converted.some(({ wrapped }) => wrapped)
					? Math.max(...converted.map(({ right }) => right)) - box.x
					: undefined,
			},
		});

		for (const { element, entry } of links) {
			linkHotspots(context, element, entry);
		}
	}

	return rasterized;
}

/** A list item that lays out apart from its list's text, as one list paragraph with its marker. */
export function markedItem(
	context: ExtractContext,
	item: Element,
	entry: TextRuns,
) {
	const laidOut = context.boxOf(item);
	const list = item.parentElement;
	if (!laidOut || !list) {
		return;
	}

	const { box, scale } = geometryOf(laidOut);
	const style = context.styleOf(item);
	const items = itemsOf(list).filter((sibling) => context.boxOf(sibling));
	const { baselines, left, right, runs, wrapLimit, wrapRight } =
		runsToParagraph(context, entry, item);
	const paragraphs = [
		{
			baselines,
			bullet: true,
			indent: left - box.x,
			level: 0,
			marker: style["list-style-type"],
			markerColor: markerColorOf(context, item),
			markerFace: markerFaceOf(context, item),
			number: ordinalsOf(list, items)[items.indexOf(item)],
			runs,
			visible: true,
		},
	];
	if (isInvisible(paragraphs)) {
		return;
	}

	context.push(item, {
		box,
		key: `${pathOf(item)}#list`,
		kind: "list",
		paragraphs,
		style: {
			align: alignOf(context, item),
			available: availableSpanOf(context, item),
			lineHeight: lineHeightOf(style) * scale,
			nowrap: false,
			rtl: isRightToLeft(context, item),
			wrapLimit: wrapLimit === undefined ? undefined : wrapLimit + left - box.x,
			wrapWidth: wrapRight === undefined ? undefined : right - box.x,
		},
	});
}

export function ordinalsOf(list: Element, items: Element[]) {
	const reversed = list.hasAttribute("reversed");
	const start = Number(list.getAttribute("start") ?? NaN);
	let ordinal = Number.isInteger(start) ? start : reversed ? items.length : 1;
	return items.map((item) => {
		const value = Number(item.getAttribute("value") ?? NaN);
		const current = Number.isInteger(value) ? value : ordinal;
		ordinal = current + (reversed ? -1 : 1);
		return current;
	});
}

/** A list's items in the columns they lay out in (CSS columns, which the layout balances), each column's items with its left and right edges. */
function columnsOf(context: ExtractContext, list: Element) {
	const balanced = [list, list.parentElement].some(
		(element) =>
			element && Number(context.styleOf(element)["column-count"]) > 1,
	);
	const columns: { items: Element[]; left: number; right: number }[] = [];
	let previousTop = -Infinity;
	for (const item of itemsOf(list)) {
		const laidOut = context.boxOf(item);
		if (!laidOut) {
			continue;
		}

		const { box } = geometryOf(laidOut);
		const column = columns.at(-1);
		if (column && !(balanced && box.y < previousTop - 1)) {
			column.items.push(item);
			column.left = Math.min(column.left, box.x);
			column.right = Math.max(column.right, box.x + box.width);
		} else {
			columns.push({ items: [item], left: box.x, right: box.x + box.width });
		}

		previousTop = box.y;
	}

	return columns;
}

/** The narrowest of a list's paragraphs' wrap limits, if any wrapped. */
function limitOf(limits: (number | undefined)[]) {
	const known = limits.filter((limit) => limit !== undefined);
	return known.length ? Math.min(...known) : undefined;
}

/** The color a list item's marker shows in, from its ::marker's style (which inherits the item's). */
function markerColorOf(context: ExtractContext, item: Element) {
	const style = context.markers.get(item) ?? context.styleOf(item);
	return parseColor(style.color)?.hex;
}

/** The font and size a list item's marker shows in, from its ::marker's style (which inherits the item's). */
function markerFaceOf(context: ExtractContext, item: Element): MarkerFace {
	const style = context.markers.get(item) ?? context.styleOf(item);
	const { font } = context.fonts.resolve(
		style["font-family"],
		Number(style["font-weight"]),
		style["font-style"] === "italic",
	);
	return {
		font,
		size: parseFloat(style["font-size"]) * context.scaleAt(item),
	};
}

const itemsOf = (list: Element) =>
	[...list.children].filter((child) => child.localName === "li");

/** A list's items and its nested lists' items, in document order. */
const itemsWithin = (context: ExtractContext, list: Element): Element[] =>
	itemsOf(list).flatMap((item) => [
		item,
		...[...item.children]
			.filter((child) => isTextList(context, child))
			.flatMap((child) => itemsWithin(context, child)),
	]);

function itemBlocks(context: ExtractContext, item: Element): ItemBlock[] {
	const own = context.ownRunsOf(item);
	if (own) {
		return [{ entry: own, kind: "text", y: own.box.y }];
	}

	return [
		...context
			.looseRunsOf(item)
			.map((entry): ItemBlock => ({ entry, kind: "text", y: entry.box.y })),
		...[...item.children].flatMap((child): ItemBlock[] => {
			const box = context.boxOf(child);
			const entry = context.ownRunsOf(child);
			return isTextList(context, child) && box
				? [{ kind: "list", list: child, y: box.y }]
				: entry && !context.drawsApart(child)
					? [{ entry, kind: "text", y: entry.box.y }]
					: [];
		}),
	].toSorted((first, second) => first.y - second.y);
}

function listParagraphs(
	context: ExtractContext,
	list: Element,
	level: number,
	textLeft: number,
	opacity: number,
	rasterized: Element[],
	/** Collects each item's text, to lay link hotspots over once the list is placed. */
	links: { element: Element; entry: TextRuns }[],
	shown = itemsOf(list),
): {
	paragraph: ListParagraph;
	right: number;
	wrapLimit?: number;
	wrapped: boolean;
}[] {
	const items = itemsOf(list).filter((item) => context.boxOf(item));
	const numbers = ordinalsOf(list, items);
	return items.flatMap((item, index) => {
		if (!shown.includes(item) || rasterized.includes(item)) {
			return [];
		}

		const style = context.styleOf(item);
		const shared = {
			level,
			lineHeight: lineHeightOf(style) * context.scaleAt(item),
			marker: style["list-style-type"],
			markerColor: markerColorOf(context, item),
			markerFace: markerFaceOf(context, item),
			number: numbers[index],
			visible: context.opacityOf(item) >= MIN_OPACITY,
		};
		const blocks = itemBlocks(context, item);
		const startsApart = blocks[0]?.kind !== "text";
		const converted = blocks.flatMap((block, blockIndex) => {
			if (block.kind === "list") {
				return listParagraphs(
					context,
					block.list,
					level + 1,
					textLeft,
					opacity,
					rasterized,
					links,
				);
			}

			inlinePills(context, block.entry.element ?? item, block.entry);
			links.push({ element: block.entry.element ?? item, entry: block.entry });
			const { baselines, left, right, runs, wrapLimit, wrapRight } =
				runsToParagraph(
					context,
					block.entry,
					item,
					shared.visible ? context.opacityOf : () => opacity,
				);
			if (!runs.length) {
				return [];
			}

			return [
				{
					paragraph: {
						...shared,
						baselines,
						bullet: blockIndex === 0 && hasMarker(style),
						indent: left - textLeft,
						runs,
					},
					right,
					// Its lines start at its indent, so the box needs that much more width.
					wrapLimit:
						wrapLimit === undefined ? undefined : wrapLimit + left - textLeft,
					wrapped: wrapRight !== undefined,
				},
			];
		});
		if (!startsApart) {
			return converted;
		}

		const lineHeight = lineHeightOf(style) * context.scaleAt(item);
		const baseline = Math.min(
			strutBaselineOf(context, item, lineHeight),
			(converted.at(0)?.paragraph.baselines.at(0) ?? Infinity) - lineHeight,
		);
		return [
			{
				paragraph: {
					...shared,
					baselines: Number.isFinite(baseline) ? [baseline] : [],
					bullet: hasMarker(style),
					indent: 0,
					runs: [
						{
							...runStyle(context, item, opacity),
							pad: true,
							text: NO_BREAK_SPACE,
							transparency: 100,
						},
					],
				},
				right: -Infinity,
				wrapped: false,
			},
			...converted,
		];
	});
}

/** Where a list item's first line's baseline sits when no text starts it: on a line of its own font at the top of its content box. */
function strutBaselineOf(
	context: ExtractContext,
	item: Element,
	lineHeight: number,
) {
	const laidOut = context.boxOf(item);
	if (!laidOut) {
		return Infinity;
	}

	const { box, scale } = geometryOf(laidOut);
	const style = context.styleOf(item);
	const { font } = context.fonts.resolve(
		style["font-family"],
		Number(style["font-weight"]),
		style["font-style"] === "italic",
	);
	const { ascent, descent } = context.fonts.metrics(font);
	const size = parseFloat(style["font-size"]) * scale;
	return (
		box.y +
		context.contentInsets(item).top * scale +
		(lineHeight - (ascent + descent) * size) / 2 +
		ascent * size
	);
}
