import type { Matrix, Placement } from "./geometry.ts";
import type { ImageMetrics } from "./media.ts";
import type { TextContext } from "./text.ts";
import type { SlideTiming } from "./timing.ts";

import { EMU_PER_PX, multiply, parseTransform, place } from "./geometry.ts";
import { escapeAttribute, px, styleOf } from "./html.ts";
import { convertText } from "./text.ts";
import {
	attributeOf,
	childNamed,
	descendantNamed,
	descendantsNamed,
} from "./xml.ts";

export interface SlideContext extends TextContext {
	/** Copies a media file from the .pptx into the deck, returning its public URL. */
	copyMedia: (relationshipId: string) => string | undefined;
	/** The slide SVG's definitions by id, for shapes drawn as SVG that reference gradients or markers. */
	defs: Map<string, string>;
	/** An image's pixel size and resolution, by its relationship id, for tiled picture fills. */
	imageMetrics: (relationshipId: string) => ImageMetrics | undefined;
	/** The slide XML's shape elements by their cNvPr id. */
	raw: Map<string, Element>;
	timing: SlideTiming;
}

/** Geometries drawn as styled boxes, which slidev2pptx writes as native rectangles. */
const BOX_GEOMETRIES = new Set(["rect", "roundRect"]);

/** Geometries drawn as single straight lines. */
const LINE_GEOMETRIES = new Set(["line", "straightConnector1"]);

/** Converts a pptx-svg shape &lt;g> (and, for groups, its children) into absolutely positioned Slidev HTML. */
export function convertShape(
	context: SlideContext,
	shape: Element,
	matrix: Matrix,
	chain: string[],
): string {
	const type = shape.getAttribute("data-ooxml-shape-type");
	const id = shape.getAttribute("data-ooxml-nv-id") ?? "";
	const placement = place(matrix, {
		cx: numberOf(shape, "cx"),
		cy: numberOf(shape, "cy"),
		rotation: numberOf(shape, "rot") / 60_000,
		x: numberOf(shape, "x"),
		y: numberOf(shape, "y"),
	});

	let html: string;
	if (type === "group") {
		const inner = [...shape.children].find((child) => child.tagName === "g");
		// pptx-svg draws a group's children in a child space of its own, which its transform maps to the slide,
		// but that transform can't place the children's OOXML frames, as for a flipped group; those use the group's own.
		const transform = inner?.getAttribute("transform") ?? "";
		const childMatrix = multiply(
			matrix,
			groupTransform(
				descendantNamed(childNamed(context.raw.get(id), "p:grpSpPr"), "a:xfrm"),
			) ?? parseTransform(transform),
		);
		html = `<div class="pptx-group" data-shape="${escapeAttribute(id)}">${[
			...(inner?.children ?? []),
		]
			.filter((child) => child.hasAttribute("data-ooxml-shape-type"))
			.map((child) =>
				convertShape(context, child, childMatrix, [...chain, transform]),
			)
			.join("")}</div>`;
	} else if (type === "picture") {
		html = convertPicture(context, shape, id, placement, chain);
	} else {
		html = convertAutoshape(context, shape, id, placement, chain);
	}

	return context.withVisibility(html, context.timing.shapes.get(id));
}

function boxOf(placement: Placement) {
	return {
		height: px(placement.height),
		left: px(placement.x),
		top: px(placement.y),
		transform: placement.rotation
			? `rotate(${String(placement.rotation)}deg)`
			: undefined,
		width: px(placement.width),
	};
}

function convertAutoshape(
	context: SlideContext,
	shape: Element,
	id: string,
	placement: Placement,
	chain: string[],
) {
	const raw = context.raw.get(id);
	const geometry = shape.getAttribute("data-ooxml-geom") ?? "rect";
	// Some geometries, like braces, draw an unpainted fill path before their stroke.
	const shapes = [...shape.children].filter((child) =>
		["ellipse", "line", "path", "polygon", "polyline", "rect"].includes(
			child.tagName,
		),
	);
	const drawn =
		shapes.find(
			(child) =>
				(child.getAttribute("fill") ?? "none") !== "none" ||
				((child.getAttribute("stroke") ?? "none") !== "none" &&
					Number(child.getAttribute("stroke-width") ?? 0) > 0),
		) ?? shapes.at(0);
	const fill = drawn?.getAttribute("fill") ?? "none";
	const stroke = drawn?.getAttribute("stroke") ?? "none";
	const strokeWidth = Number(drawn?.getAttribute("stroke-width") ?? 0);
	const effects = descendantNamed(childNamed(raw, "p:spPr"), "a:effectLst");
	if (effects?.children.length) {
		context.lose(
			`Dropped ${[...effects.children].map((effect) => effect.tagName.replace("a:", "")).join(", ")} on shape ${id}`,
		);
	}

	const parts: string[] = [];
	const visible = fill !== "none" || (stroke !== "none" && strokeWidth > 0);
	if (visible) {
		const flipped =
			!!shape.getAttribute("data-ooxml-flip-h") ||
			!!shape.getAttribute("data-ooxml-flip-v");
		if (
			BOX_GEOMETRIES.has(geometry) &&
			drawn?.tagName === "rect" &&
			!fill.startsWith("url(") &&
			!drawn.getAttribute("stroke-dasharray")
		) {
			parts.push(drawBox(drawn, placement, fill, stroke, strokeWidth));
		} else if (
			LINE_GEOMETRIES.has(geometry) &&
			!drawn?.getAttribute("marker-start") &&
			!drawn?.getAttribute("marker-end") &&
			!flipped &&
			stroke !== "none"
		) {
			parts.push(drawLine(drawn, placement, stroke, strokeWidth));
		} else if (geometry === "ellipse" && !fill.startsWith("url(")) {
			parts.push(drawEllipse(drawn, placement, fill, stroke, strokeWidth));
		} else {
			parts.push(drawSvg(context, shape, placement, chain, strokeWidth));
		}
	}

	const body = descendantNamed(raw, "p:txBody");
	const rendered = [...shape.children].find(
		(child) => child.tagName === "text",
	);
	// PowerPoint shrinks text by the scale its autofit recorded, which pptx-svg reports only sometimes.
	// That's only trusted for sizes the slide sets, as pptx-svg can inherit others from the wrong master placeholder.
	const reported = Number(shape.getAttribute("data-ooxml-font-scale") ?? -1);
	const recorded = descendantsNamed(body, "a:r").every((run) =>
		childNamed(run, "a:rPr")?.hasAttribute("sz"),
	)
		? Number(
				childNamed(childNamed(body, "a:bodyPr"), "a:normAutofit")?.getAttribute(
					"fontScale",
				) ?? -1,
			)
		: -1;
	const fontScale = reported > 0 ? reported : recorded;
	const box = [...shape.children].find((child) => child.tagName === "rect");
	// pptx-svg writes -1 for an unset inset, which is PowerPoint's default 0.1".
	const declaredInset = Number(
		childNamed(body, "a:bodyPr")?.getAttribute("lIns") ??
			attributeOf(shape, "data-ooxml-l-ins") ??
			-1,
	);
	const leftInset = (declaredInset < 0 ? 91_440 : declaredInset) / EMU_PER_PX;
	const text = convertText(
		context,
		body,
		rendered,
		context.timing.paragraphs.get(id),
		fontScale > 0 ? fontScale / 100_000 : 0,
		Number(box?.getAttribute("x") ?? 0) + leftInset,
	);
	if (text) {
		parts.push(drawText(shape, body, rendered, placement, text));
	}

	return parts.join("");
}

function convertPicture(
	context: SlideContext,
	shape: Element,
	id: string,
	placement: Placement,
	chain: string[],
) {
	const raw = context.raw.get(id);
	const blip = descendantNamed(raw, "a:blip");
	const svgBlip = descendantNamed(blip, "asvg:svgBlip");
	const relationship =
		svgBlip?.getAttribute("r:embed") ?? blip?.getAttribute("r:embed");
	const source = relationship ? context.copyMedia(relationship) : undefined;
	if (!source) {
		context.lose(`Drew picture ${id} as SVG, as its image couldn't be copied`);
		return drawSvg(context, shape, placement, chain, 0);
	}

	if (
		descendantNamed(raw, "a:videoFile") ||
		descendantNamed(raw, "p14:media")
	) {
		context.lose(`Kept only the poster image of video ${id}`);
	}

	const alt = descendantNamed(raw, "p:cNvPr")?.getAttribute("descr") ?? "";
	const crop = descendantNamed(raw, "a:srcRect");
	const [left, top, right, bottom] = ["l", "t", "r", "b"].map(
		(side) => Number(crop?.getAttribute(side) ?? 0) / 100_000,
	);
	const alpha = descendantNamed(blip, "a:alphaModFix")?.getAttribute("amt");
	const opacity = alpha ? Number(alpha) / 100_000 : undefined;
	const flipH =
		!!shape.getAttribute("data-ooxml-flip-h") !== placement.mirrored;
	const flipV = !!shape.getAttribute("data-ooxml-flip-v");
	const flips = [flipH ? "scaleX(-1)" : "", flipV ? "scaleY(-1)" : ""]
		.filter(Boolean)
		.join(" ");
	const box = boxOf(placement);
	const transform =
		[box.transform, flips].filter(Boolean).join(" ") || undefined;
	const geometry =
		descendantNamed(childNamed(raw, "p:spPr"), "a:prstGeom")?.getAttribute(
			"prst",
		) ?? shape.getAttribute("data-ooxml-geom");
	const radius = geometry === "ellipse" ? "50%" : undefined;
	if (geometry && geometry !== "rect" && geometry !== "ellipse") {
		context.lose(`Drew picture ${id}'s ${geometry} shape as a rectangle`);
	}

	if (!left && !top && !right && !bottom) {
		return `<img class="pptx-shape" src="${escapeAttribute(source)}" alt="${escapeAttribute(alt)}" style="${styleOf(
			{
				...box,
				"border-radius": radius,
				opacity,
				transform,
			},
		)}">`;
	}

	const width = placement.width / (1 - left - right);
	const height = placement.height / (1 - top - bottom);
	return `<div class="pptx-shape pptx-crop" style="${styleOf({
		...box,
		"border-radius": radius,
		opacity,
		transform,
	})}"><img src="${escapeAttribute(source)}" alt="${escapeAttribute(alt)}" style="${styleOf(
		{
			height: px(height),
			left: px(-left * width),
			position: "absolute",
			top: px(-top * height),
			width: px(width),
		},
	)}"></div>`;
}

function dashOf(drawn: Element | undefined) {
	const dashes = drawn?.getAttribute("stroke-dasharray");
	if (!dashes) {
		return "solid";
	}

	const [first] = dashes.split(/[\s,]+/).map(Number);
	const width = Number(drawn?.getAttribute("stroke-width") ?? 1);
	return first <= width * 1.5 ? "dotted" : "dashed";
}

function drawBox(
	drawn: Element,
	placement: Placement,
	fill: string,
	stroke: string,
	strokeWidth: number,
) {
	const hasStroke = stroke !== "none" && strokeWidth > 0;
	// PowerPoint centers strokes on the shape's edge; CSS borders sit inside the box.
	const grow = hasStroke ? strokeWidth / 2 : 0;
	const radius = Number(drawn.getAttribute("rx") ?? 0);
	return `<div class="pptx-shape" style="${styleOf({
		...boxOf({
			...placement,
			height: placement.height + grow * 2,
			width: placement.width + grow * 2,
			x: placement.x - grow,
			y: placement.y - grow,
		}),
		"background-color": fill === "none" ? undefined : fill,
		border: hasStroke
			? `${px(strokeWidth)} ${dashOf(drawn)} ${stroke}`
			: undefined,
		"border-radius": radius ? px(radius + grow) : undefined,
		opacity: opacityOf(drawn),
	})}"></div>`;
}

function drawEllipse(
	drawn: Element | undefined,
	placement: Placement,
	fill: string,
	stroke: string,
	strokeWidth: number,
) {
	const { height, width } = placement;
	return `<svg class="pptx-shape" viewBox="0 0 ${String(width)} ${String(height)}" style="${styleOf(
		{
			...boxOf(placement),
			overflow: "visible",
		},
	)}"><ellipse cx="${String(width / 2)}" cy="${String(height / 2)}" rx="${String(width / 2)}" ry="${String(height / 2)}" fill="${fill}"${stroke !== "none" ? ` stroke="${stroke}" stroke-width="${String(strokeWidth)}"` : ""}${drawn?.getAttribute("fill-opacity") ? ` fill-opacity="${drawn.getAttribute("fill-opacity") ?? ""}"` : ""}/></svg>`;
}

function drawLine(
	drawn: Element | undefined,
	placement: Placement,
	stroke: string,
	strokeWidth: number,
) {
	// A line's frame is its bounding box: it runs corner to corner, then rotates with the frame.
	const length = Math.hypot(placement.width, placement.height);
	const angle =
		(Math.atan2(placement.height, placement.width) * 180) / Math.PI +
		placement.rotation;
	const centerX = placement.x + placement.width / 2;
	const centerY = placement.y + placement.height / 2;
	return `<div class="pptx-shape" style="${styleOf({
		"border-top": `${px(strokeWidth)} ${dashOf(drawn)} ${stroke}`,
		height: "0px",
		left: px(centerX - length / 2),
		opacity: opacityOf(drawn, "stroke-opacity"),
		top: px(centerY - strokeWidth / 2),
		transform: angle
			? `rotate(${String(Math.round(angle * 100) / 100)}deg)`
			: undefined,
		width: px(length),
	})}"></div>`;
}

/** Draws a shape's pptx-svg rendering, minus its text, as an inline SVG around the shape. */
function drawSvg(
	context: SlideContext,
	shape: Element,
	placement: Placement,
	chain: string[],
	strokeWidth: number,
) {
	// Leave room for strokes, arrowheads, and rotation.
	const margin = strokeWidth * 4 + 8;
	const radius = Math.hypot(placement.width, placement.height) / 2;
	const centerX = placement.x + placement.width / 2;
	const centerY = placement.y + placement.height / 2;
	const halfWidth =
		(placement.rotation ? radius : placement.width / 2) + margin;
	const halfHeight =
		(placement.rotation ? radius : placement.height / 2) + margin;
	const left = centerX - halfWidth;
	const top = centerY - halfHeight;
	const retile = tilesFor(context, shape);
	const content = retile(
		[...shape.children]
			.filter((child) => child.tagName !== "text")
			.map((child) => child.outerHTML)
			.join(""),
	);
	const opened = chain
		.map((transform) => `<g transform="${escapeAttribute(transform)}">`)
		.join("");
	return `<svg class="pptx-shape" xmlns="http://www.w3.org/2000/svg" viewBox="${[left, top, halfWidth * 2, halfHeight * 2].map((value) => String(Math.round(value * 100) / 100)).join(" ")}" style="${styleOf(
		{
			height: px(halfHeight * 2),
			left: px(left),
			top: px(top),
			width: px(halfWidth * 2),
		},
	)}">${retile(definitionsFor(context.defs, content))}${opened}${content}${"</g>".repeat(chain.length)}</svg>`;
}

/**
 * Resizes a tiled picture fill's patterns in SVG markup as PowerPoint tiles it:
 * pptx-svg sizes tiles without the image's resolution, where PowerPoint scales the image's
 * size in inches, then aligns the first tile in the shape's box and offsets it.
 */
function tilesFor(context: SlideContext, shape: Element) {
	const raw = context.raw.get(shape.getAttribute("data-ooxml-nv-id") ?? "");
	const blipFill = childNamed(childNamed(raw, "p:spPr"), "a:blipFill");
	const tile = childNamed(blipFill, "a:tile");
	const metrics = context.imageMetrics(
		childNamed(blipFill, "a:blip")?.getAttribute("r:embed") ?? "",
	);
	if (!tile || !metrics) {
		return (markup: string) => markup;
	}

	if ((tile.getAttribute("flip") ?? "none") !== "none") {
		context.lose(
			`Tiled the picture fill of shape ${shape.getAttribute("data-ooxml-nv-id") ?? ""} without mirroring its tiles`,
		);
	}

	const attribute = (name: string, fallback: number) =>
		Number(tile.getAttribute(name) ?? fallback);
	const width =
		(metrics.width / metrics.dpiX) * 96 * (attribute("sx", 100_000) / 100_000);
	const height =
		(metrics.height / metrics.dpiY) * 96 * (attribute("sy", 100_000) / 100_000);
	const box = {
		height: numberOf(shape, "cy") / EMU_PER_PX,
		width: numberOf(shape, "cx") / EMU_PER_PX,
		x: numberOf(shape, "x") / EMU_PER_PX,
		y: numberOf(shape, "y") / EMU_PER_PX,
	};
	const align = tile.getAttribute("algn") ?? "tl";
	const share = (start: string, end: string) =>
		align.includes(end) ? 1 : align.includes(start) ? 0 : 0.5;
	const round = (value: number) => String(Math.round(value * 100) / 100);
	const sized = {
		height: round(height),
		width: round(width),
		x: round(
			box.x +
				share("l", "r") * (box.width - width) +
				attribute("tx", 0) / EMU_PER_PX,
		),
		y: round(
			box.y +
				share("t", "b") * (box.height - height) +
				attribute("ty", 0) / EMU_PER_PX,
		),
	};

	return (markup: string) =>
		markup.replaceAll(/<pattern\b[^>]*>\s*<image\b[^>]*>/g, (opening) => {
			const image = opening.indexOf("<image");
			return opening.replaceAll(
				/\b(x|y|width|height)="[^"]*"/g,
				(pair, name: keyof typeof sized, offset: number) =>
					// The image keeps its place in its tile; only its size changes.
					offset > image && (name === "x" || name === "y")
						? pair
						: `${name}="${sized[name]}"`,
			);
		});
}

/** The &lt;defs> SVG markup needs: only those it references, and those they reference, as some (like a background's picture fill) are huge. */
function definitionsFor(defs: Map<string, string>, markup: string) {
	const used = new Map<string, string>();
	const queue = [markup];
	for (let next = queue.pop(); next !== undefined; next = queue.pop()) {
		for (const [, id] of next.matchAll(
			/(?:url\(\s*['"]?|href=")#([^'")\s]+)/g,
		)) {
			const definition = defs.get(id);
			if (definition && !used.has(id)) {
				used.set(id, definition);
				queue.push(definition);
			}
		}
	}

	// A shape's own definitions are already in its markup.
	const missing = [...used]
		.filter(([id]) => !markup.includes(` id="${id}"`))
		.map(([, definition]) => definition);
	return missing.length ? `<defs>${missing.join("")}</defs>` : "";
}

function drawText(
	shape: Element,
	body: Element | undefined,
	rendered: Element | undefined,
	placement: Placement,
	html: string,
) {
	const properties = childNamed(body, "a:bodyPr");
	const inset = (side: string, fallback: number) => {
		const value =
			properties?.getAttribute(`${side}Ins`) ??
			shape.getAttribute(`data-ooxml-${side[0]}-ins`);
		return (
			Number(value === null || Number(value) < 0 ? fallback : value) /
			EMU_PER_PX
		);
	};

	const anchor =
		properties?.getAttribute("anchor") ??
		attributeOf(shape, "data-ooxml-anchor") ??
		inferAnchor(rendered, placement);
	const wrap =
		(properties?.getAttribute("wrap") ??
			shape.getAttribute("data-ooxml-wrap")) !== "none";
	return `<div class="pptx-shape pptx-text" style="${styleOf({
		...boxOf(placement),
		"justify-content":
			anchor === "ctr" ? "center" : anchor === "b" ? "flex-end" : undefined,
		"padding-bottom": px(inset("b", 45_720)),
		"padding-left": px(inset("l", 91_440)),
		"padding-right": px(inset("r", 91_440)),
		"padding-top": px(inset("t", 45_720)),
		"white-space": wrap ? undefined : "pre",
	})}">${html}</div>`;
}

/** Infers a text box's vertical anchor from where pptx-svg laid out its lines, for anchors inherited from layouts. */
function inferAnchor(rendered: Element | undefined, placement: Placement) {
	const ys = [...(rendered?.children ?? [])]
		.map((line) => Number(line.getAttribute("y")))
		.filter((y) => Number.isFinite(y) && y > 0);
	if (!ys.length) {
		return "t";
	}

	const top = Math.min(...ys) - placement.y;
	const bottom = placement.y + placement.height - Math.max(...ys);
	if (Math.abs(top - bottom) < placement.height * 0.15) {
		return "ctr";
	}

	return bottom < top ? "b" : "t";
}

function numberOf(shape: Element, name: string) {
	return Number(shape.getAttribute(`data-ooxml-${name}`) ?? 0);
}

function opacityOf(drawn: Element | undefined, attribute = "fill-opacity") {
	const value = drawn?.getAttribute(attribute);
	return value && Number(value) < 1 ? Number(value) : undefined;
}

/**
 * A group's transform from its child space to its parent's, in px: its children's frame
 * (chOff, chExt) scaled into its own (off, ext), then flipped and turned about its center,
 * as PowerPoint flips a shape before turning it.
 */
function groupTransform(xfrm: Element | undefined): Matrix | undefined {
	const frame = (name: string, x: string, y: string) => {
		const element = childNamed(xfrm, name);
		return element
			? {
					x: Number(element.getAttribute(x) ?? 0) / EMU_PER_PX,
					y: Number(element.getAttribute(y) ?? 0) / EMU_PER_PX,
				}
			: undefined;
	};

	const offset = frame("a:off", "x", "y");
	const extent = frame("a:ext", "cx", "cy");
	if (!xfrm || !offset || !extent) {
		return undefined;
	}

	const childOffset = frame("a:chOff", "x", "y") ?? offset;
	const childExtent = frame("a:chExt", "cx", "cy") ?? extent;
	const scale = {
		x: childExtent.x ? extent.x / childExtent.x : 1,
		y: childExtent.y ? extent.y / childExtent.y : 1,
	};
	const center = { x: offset.x + extent.x / 2, y: offset.y + extent.y / 2 };
	const radians =
		((Number(xfrm.getAttribute("rot") ?? 0) / 60_000) * Math.PI) / 180;
	const cos = Math.cos(radians);
	const sin = Math.sin(radians);
	const flipX = xfrm.getAttribute("flipH") === "1" ? -1 : 1;
	const flipY = xfrm.getAttribute("flipV") === "1" ? -1 : 1;
	const translate = (x: number, y: number): Matrix => ({
		a: 1,
		b: 0,
		c: 0,
		d: 1,
		e: x,
		f: y,
	});
	return [
		translate(center.x, center.y),
		{ a: cos, b: sin, c: -sin, d: cos, e: 0, f: 0 },
		{ a: flipX, b: 0, c: 0, d: flipY, e: 0, f: 0 },
		translate(-center.x, -center.y),
		translate(offset.x, offset.y),
		{ a: scale.x, b: 0, c: 0, d: scale.y, e: 0, f: 0 },
		translate(-childOffset.x, -childOffset.y),
	].reduce(multiply);
}
