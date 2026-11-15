import type { ExtractOptions } from "./extract/context.ts";
import type { LaidOutBox, Layout } from "./layout.ts";

import { boxStyle, paintItems } from "./extract/boxes.ts";
import {
	createExtractContext,
	isMath,
	MIN_OPACITY,
} from "./extract/context.ts";
import { geometryOf } from "./extract/geometry.ts";
import { imageItem, rasterItem } from "./extract/images.ts";
import {
	isListAside,
	isMarkedApart,
	isTextList,
	listItem,
	markedItem,
} from "./extract/lists.ts";
import { svgItems } from "./extract/svg.ts";
import {
	codeItem,
	inlinePills,
	linkHotspots,
	textItem,
} from "./extract/text.ts";
import { isBackgroundOnly } from "./layout.ts";

/** Elements whose descendants would draw more shapes than this render as one image instead. */
const MAX_NATIVE_SHAPES = 400;

/** Elements that render only in a browser, which PowerPoint can't show. */
const UNSUPPORTED_EMBEDS = [
	[".mermaid", "a Mermaid diagram"],
	[".slidev-tweet", "an embedded post"],
	["iframe", "an embedded page"],
	["video", "a video"],
] as const;

/** Turns a laid-out slide state into model items, and descriptions of what it left out. */
export function extractState(layout: Layout, options: ExtractOptions) {
	const context = createExtractContext(layout, options);
	const { styleOf } = context;

	const shapes = new Map<Element, number>();

	/**
	 * Roughly how many shapes an element's descendants draw. Plain inline runs
	 * join their paragraph's text, and a run of abutting siblings with the same
	 * background, like highlighted code tokens, shares one box; what isn't
	 * displayed draws nothing; anything else may be a shape of its own.
	 */
	const shapesUnder = (element: Element): number => {
		let count = shapes.get(element);
		if (count !== undefined) {
			return count;
		}

		count = 0;
		let previous: string | undefined;
		for (const child of element.children) {
			const style = styleOf(child);
			if (style.display === "none") {
				previous = undefined;
				continue;
			}

			const inline = style.display === "inline" && !child.closest("svg");
			const painted = inline && boxStyle(context, child, 1);
			const background = painted ? style["background-color"] : undefined;
			if (!inline || (painted && background !== previous)) {
				count += 1;
			}

			previous = background;
			count += shapesUnder(child);
		}

		shapes.set(element, count);
		return count;
	};

	const isDense = (element: Element) =>
		shapesUnder(element) > MAX_NATIVE_SHAPES;

	const isClippedAway = (element: Element, laidOut: LaidOutBox) => {
		const style = styleOf(element);
		const clips = ["clip", "overflow", "overflow-x", "overflow-y"].some(
			(property) => /\b(?:hidden|clip|rect)\b/.test(style[property] ?? ""),
		);
		return (
			/^inset\(\s*50%\s*\)$/.test(style["clip-path"] ?? "") ||
			(clips &&
				(laidOut.width <= 1 || laidOut.height <= 1) &&
				element.textContent.trim() !== "")
		);
	};

	const walkApart = (element: Element, drawsApart = context.drawsApart) => {
		for (const child of element.children) {
			if (drawsApart(child)) {
				walk(child);
			} else {
				walkApart(child, drawsApart);
			}
		}
	};

	const walk = (element: Element) => {
		if (isBackgroundOnly(element, layout.styles)) {
			return;
		}

		const opacity = context.opacityOf(element);
		if (opacity < MIN_OPACITY) {
			if (styleOf(element).visibility === "hidden") {
				for (const child of element.children) {
					walk(child);
				}
			}

			return;
		}

		const embed = UNSUPPORTED_EMBEDS.find(([selector]) =>
			element.matches(selector),
		);
		if (embed) {
			const source = element.getAttribute("src");
			context.problems.push(
				`Couldn't export ${embed[1]}${source ? ` (${source})` : ""}.`,
			);
			return;
		}

		if (isMath(element)) {
			rasterItem(
				context,
				element,
				element.querySelector("annotation")?.textContent.trim(),
			);
			return;
		}

		const laidOut = context.boxOf(element);
		if (!laidOut) {
			if (context.isInline(element)) {
				walkApart(element);
			}

			return;
		}

		if (isClippedAway(element, laidOut)) {
			return;
		}

		const geometry = geometryOf(laidOut);
		const paint = paintItems(context, element, opacity, geometry);
		if (!paint) {
			rasterItem(context, element);
			return;
		}

		context.push(element, ...paint);

		if (element.localName === "img") {
			imageItem(context, element, opacity, geometry);
		} else if (element.localName === "svg") {
			svgItems(context, element);
		} else if (element.localName === "pre") {
			codeItem(context, element);
		} else if (
			isDense(element) &&
			[...element.children].filter(isDense).length !== 1
		) {
			rasterItem(context, element);
		} else if (isTextList(context, element)) {
			const rasterized = listItem(context, element, opacity);
			walkApart(
				element,
				(child) =>
					!rasterized.some((raster) => raster.contains(child)) &&
					(context.drawsApart(child) || isListAside(context, child)),
			);
		} else {
			const entry = context.ownRunsOf(element);
			if (entry) {
				inlinePills(context, element, entry);
				if (isMarkedApart(context, element)) {
					markedItem(context, element, entry);
				} else {
					textItem(context, element, entry);
				}

				linkHotspots(context, element, entry);

				walkApart(element);
				return;
			}

			for (const child of element.children) {
				walk(child);
			}

			for (const loose of context.looseRunsOf(element)) {
				inlinePills(context, element, loose);
				textItem(context, element, loose);
				linkHotspots(context, element, loose);
			}
		}
	};

	for (const layer of layout.document.querySelector("#slide-content")
		?.children ?? []) {
		walk(layer);
	}

	return { items: context.items, problems: context.problems };
}
