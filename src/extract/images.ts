import path from "node:path";

import type { Box, Canvas } from "../model.ts";
import type { ExtractContext } from "./context.ts";
import type { Geometry } from "./geometry.ts";

import { imageSize, withoutQuery } from "../assets.ts";
import { transparencyOf } from "../color.ts";
import { ID } from "../layout.ts";
import { boxShadowOf, scaledShadow } from "./boxes.ts";
import { pathOf } from "./context.ts";
import {
	centeredAt,
	centerOf,
	clipBox,
	fitSize,
	geometryOf,
	insetBox,
	intersect,
	outOfFrame,
	positionIn,
	sameBox,
} from "./geometry.ts";

/** A filter's blur length is a standard deviation; the blur shows to about this many of them. */
const BLUR_REACH = 3;

/** Room around a raster's element for strokes and antialiasing, beyond its shadows' reach. */
const RASTER_MARGIN = 4;

export function fileNameOf(src: string) {
	const name = path.posix.basename(withoutQuery(src));
	try {
		return decodeURIComponent(name);
	} catch {
		return name;
	}
}

export function imageItem(
	context: ExtractContext,
	image: Element,
	opacity: number,
	{ box: border, flip, rotation, scale }: Geometry,
) {
	const src = image.getAttribute("src");
	if (!src) {
		return;
	}

	const data = context.assets.get(src);
	const file = context.assets.save(src);
	if (!data || !file) {
		context.problems.push(`Couldn't load the image ${src}.`);
		return;
	}

	const style = context.styleOf(image);
	const box = insetBox(border, context.contentInsets(image), scale);
	const natural = imageSize(data);
	const drawn = natural
		? positionIn(
				box,
				objectSize(style["object-fit"], box, natural, scale),
				style["object-position"],
				(length, free) => context.resolve(image, length, free),
			)
		: box;
	// PowerPoint crops a picture before flipping and turning it, so a transformed one is cropped in its own frame, by the clips turned with it.
	const frame =
		rotation || flip ? { center: centerOf(box), flip, rotation } : undefined;
	const clip = frame && context.clipOf(image, false, frame);
	const content = clip
		? clipBox(intersect(box, drawn), clip)
		: intersect(box, drawn);
	if (frame && (content.width <= 0 || content.height <= 0)) {
		return;
	}

	const placed = frame
		? centeredAt(content, outOfFrame(frame, centerOf(content)))
		: content;
	const radius = context.radiusOf(image) * scale;
	const shadow = boxShadowOf(context, image, opacity);
	context.push(image, {
		alt: image.getAttribute("alt") ?? "",
		box: placed,
		crop: sameBox(content, drawn)
			? undefined
			: {
					...drawn,
					x: placed.x + drawn.x - content.x,
					y: placed.y + drawn.y - content.y,
				},
		file,
		flip: flip || undefined,
		key: `${pathOf(image)}#img:${src}`,
		kind: "image",
		label: src.startsWith("data:") ? context.labelOf(image) : fileNameOf(src),
		radius,
		rotation: rotation || undefined,
		shadow: shadow && scaledShadow(shadow, scale),
		src,
		transparency: transparencyOf(1, opacity),
	});
}

export function rasterItem(
	context: ExtractContext,
	element: Element,
	alt = context.textOf(element),
) {
	const box = rasterBox(context, element);
	if (box.width <= 0 || box.height <= 0) {
		return;
	}

	context.push(element, {
		alt,
		box,
		key: `${pathOf(element)}#raster`,
		kind: "raster",
		label: context.labelOf(element),
		selector: `[${ID}="${context.idOf(element)}"]`,
	});
}

function objectSize(fit: string, box: Box, natural: Canvas, scale: number) {
	const contain = fitSize("contain", box, natural.width / natural.height);
	const own = { height: natural.height * scale, width: natural.width * scale };
	const sizes: Partial<Record<string, Canvas>> = {
		contain,
		cover: fitSize("cover", box, natural.width / natural.height),
		none: own,
		"scale-down": own.width > contain.width ? contain : own,
	};
	return sizes[fit] ?? box;
}

/** How far shadows and filters may paint past an element or its descendants: at most the sum of their lengths. */
function paintOverflowOf(context: ExtractContext, element: Element) {
	return Math.max(
		0,
		...[element, ...element.querySelectorAll("*")].map((node) => {
			const style = context.styleOf(node);
			return Math.max(
				...["box-shadow", "filter", "text-shadow"].map(
					(property) =>
						[
							...(style[property] ?? "").matchAll(/-?[\d.]+(?:px|r?em)\b/g),
						].reduce(
							(sum, [length]) => sum + Math.abs(context.resolve(node, length)),
							0,
						) * (property === "filter" ? BLUR_REACH : 1),
				),
			);
		}),
	);
}

/** The whole-pixel area covering an element and everything inside it, or the whole slide when anything is rotated. */
function rasterBox(context: ExtractContext, element: Element) {
	const { canvas } = context;
	const slide = { height: canvas.height, width: canvas.width, x: 0, y: 0 };
	let left = Infinity;
	let top = Infinity;
	let right = -Infinity;
	let bottom = -Infinity;
	for (const node of [element, ...element.querySelectorAll("*")]) {
		const laidOut = context.boxOf(node);
		if (!laidOut) {
			continue;
		}

		const { box, rotation } = geometryOf(laidOut);
		if (rotation) {
			return slide;
		}

		left = Math.min(left, box.x);
		top = Math.min(top, box.y);
		right = Math.max(right, box.x + box.width);
		bottom = Math.max(bottom, box.y + box.height);
	}

	if (!Number.isFinite(left)) {
		return slide;
	}

	const margin = RASTER_MARGIN + paintOverflowOf(context, element);
	const x = Math.max(0, Math.floor(left - margin));
	const y = Math.max(0, Math.floor(top - margin));
	return {
		height: Math.min(canvas.height, Math.ceil(bottom + margin)) - y,
		width: Math.min(canvas.width, Math.ceil(right + margin)) - x,
		x,
		y,
	};
}
