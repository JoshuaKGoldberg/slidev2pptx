import type JSZip from "jszip";

import {
	childNamed,
	descendantNamed,
	descendantsNamed,
	parseXml,
} from "./xml.ts";

/** Theme font references' prefixes, by a font reference's index. */
const THEME_FONTS: Partial<Record<string, string>> = {
	major: "mj",
	minor: "mn",
};

/** Parts with shape trees, which SmartArt can be in. */
const SHAPE_TREE_PARTS =
	/^ppt\/(?:slides|slideLayouts|slideMasters)\/[^/]+\.xml$/;

/**
 * Replaces each SmartArt diagram in a deck's slides, layouts, and masters with
 * the shapes PowerPoint last drew it as, from its drawing part, as pptx-svg
 * draws nothing for diagrams. A diagram becomes a group with its frame's id,
 * so its animations still apply.
 * @returns What each part it changed lost, by part name.
 */
export async function inlineDiagrams(
	zip: JSZip,
	relationshipsOf: (part: string) => Promise<Map<string, string>>,
) {
	const changed = new Map<string, string[]>();
	for (const part of Object.keys(zip.files).filter((name) =>
		SHAPE_TREE_PARTS.test(name),
	)) {
		const xml = (await zip.file(part)?.async("string")) ?? "";
		if (!xml.includes("<dgm:relIds")) {
			continue;
		}

		const rels = await relationshipsOf(part);
		const losses: string[] = [];
		let lastId = Math.max(
			0,
			...Array.from(xml.matchAll(/<p:cNvPr\b[^>]*?\sid="(\d+)"/g), ([, id]) =>
				Number(id),
			),
		);
		const nextId = () => String((lastId += 1));
		let inlined = "";
		let end = 0;
		for (const match of xml.matchAll(
			/<p:graphicFrame\b[\s\S]*?<\/p:graphicFrame>/g,
		)) {
			const frame = parseXml(match[0]).documentElement;
			const drawing = await drawingOf(zip, rels, frame);
			const name =
				descendantNamed(frame, "p:cNvPr")?.getAttribute("name") ?? "";
			if (drawing === undefined) {
				if (descendantNamed(frame, "dgm:relIds")) {
					losses.push(`Dropped SmartArt "${name}", as it has no drawing`);
				}

				continue;
			}

			if (/\sr:(?:embed|id|link)="/.test(drawing)) {
				losses.push(
					`Dropped pictures and links in SmartArt "${name}", as its drawing's relationships aren't carried over`,
				);
			}

			inlined += xml.slice(end, match.index) + groupOf(frame, drawing, nextId);
			end = match.index + match[0].length;
		}

		zip.file(part, inlined + xml.slice(end));
		changed.set(part, losses);
	}

	return changed;
}

/** A diagram frame's drawing part, through its data part's reference, or undefined if it isn't a diagram or has none. */
async function drawingOf(
	zip: JSZip,
	rels: Map<string, string>,
	frame: Element,
) {
	const ids = descendantNamed(frame, "dgm:relIds");
	const data = rels.get(ids?.getAttribute("r:dm") ?? "");
	if (!data) {
		return undefined;
	}

	// The data part names the drawing by a relationship of the part the diagram is in.
	const reference = /<dsp:dataModelExt\b[^>]*?\srelId="([^"]+)"/.exec(
		(await zip.file(data)?.async("string")) ?? "",
	)?.[1];
	const drawing = rels.get(reference ?? "");
	return drawing ? await zip.file(drawing)?.async("string") : undefined;
}

/**
 * A diagram frame as a group of its drawing's shapes, which are placed in the frame's box.
 * Each shape's text gets its own text box at the drawing's text transform, where PowerPoint lays it out.
 */
function groupOf(frame: Element, drawing: string, nextId: () => string) {
	// SmartArt offsets its colors by nothing, which pptx-svg rounds off by a little.
	const root = parseXml(
		drawing.replaceAll(/<a:(?:alphaOff|hueOff|lumOff|satOff) val="0"\/>/g, ""),
	).documentElement;
	// The drawing's own namespace prefixes, other than its "dsp:" one, which becomes "p:".
	const namespaces = [...root.attributes]
		.filter(({ name }) => name.startsWith("xmlns:") && name !== "xmlns:dsp")
		.map(({ name, value }) => ` ${name}="${value}"`)
		.join("");
	const transform = childNamed(frame, "p:xfrm");
	const offset = childNamed(transform, "a:off");
	const extent = childNamed(transform, "a:ext");
	const width = extent?.getAttribute("cx") ?? "0";
	const height = extent?.getAttribute("cy") ?? "0";
	return `<p:grpSp${namespaces}><p:nvGrpSpPr>${descendantNamed(frame, "p:cNvPr")?.outerHTML ?? ""}<p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm${attributesOf(transform)}><a:off x="${offset?.getAttribute("x") ?? "0"}" y="${offset?.getAttribute("y") ?? "0"}"/><a:ext cx="${width}" cy="${height}"/><a:chOff x="0" y="0"/><a:chExt cx="${width}" cy="${height}"/></a:xfrm></p:grpSpPr>${shapesOf(childNamed(root, "dsp:spTree"), nextId)}</p:grpSp>`;
}

/** A drawing's shapes and groups as slide shapes and groups. */
function shapesOf(tree: Element | undefined, nextId: () => string): string {
	return [...(tree?.children ?? [])]
		.map((child) => {
			if (child.tagName === "dsp:grpSp") {
				return `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${nextId()}" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr>${contentOf(childNamed(child, "dsp:grpSpPr"))}</p:grpSpPr>${shapesOf(child, nextId)}</p:grpSp>`;
			}

			if (child.tagName !== "dsp:sp") {
				return "";
			}

			const style = childNamed(child, "dsp:style");
			const styleXml = style ? `<p:style>${contentOf(style)}</p:style>` : "";
			const body = childNamed(child, "dsp:txBody");
			const defaults = textDefaults(style);
			const bodyXml = body
				? `<p:txBody>${[...body.children]
						.map((part) =>
							part.tagName === "a:lstStyle" && !part.children.length && defaults
								? defaults
								: part.outerHTML,
						)
						.join("")}</p:txBody>`
				: "";
			const textTransform = childNamed(child, "dsp:txXfrm");
			const shape = `<p:sp><p:nvSpPr><p:cNvPr id="${nextId()}" name=""/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>${contentOf(childNamed(child, "dsp:spPr"))}</p:spPr>${styleXml}${textTransform ? "" : bodyXml}</p:sp>`;
			if (
				!textTransform ||
				!descendantsNamed(body, "a:t").some((text) => text.textContent)
			) {
				return shape;
			}

			// The text transform can differ from the shape's text rectangle, and turns apart from the shape.
			return `${shape}<p:sp><p:nvSpPr><p:cNvPr id="${nextId()}" name=""/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm${attributesOf(textTransform)}>${contentOf(textTransform)}</a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr>${styleXml}${bodyXml}</p:sp>`;
		})
		.join("");
}

/**
 * A list style giving text its shape style's font and color, as pptx-svg
 * leaves those off text in groups, which then gets its defaults instead.
 */
function textDefaults(style: Element | undefined) {
	const reference = childNamed(style, "a:fontRef");
	const color = reference?.firstElementChild;
	const theme = THEME_FONTS[reference?.getAttribute("idx") ?? ""];
	const properties = `${color ? `<a:solidFill>${color.outerHTML}</a:solidFill>` : ""}${theme ? `<a:latin typeface="+${theme}-lt"/><a:ea typeface="+${theme}-ea"/><a:cs typeface="+${theme}-cs"/>` : ""}`;
	return properties
		? `<a:lstStyle>${Array.from(
				{ length: 9 },
				(_, index) =>
					`<a:lvl${String(index + 1)}pPr><a:defRPr>${properties}</a:defRPr></a:lvl${String(index + 1)}pPr>`,
			).join("")}</a:lstStyle>`
		: undefined;
}

/** An element's attributes, as XML to open a tag with. */
function attributesOf(element: Element | undefined) {
	return [...(element?.attributes ?? [])]
		.map(
			({ name, value }) =>
				` ${name}="${value.replaceAll("&", "&amp;").replaceAll('"', "&quot;")}"`,
		)
		.join("");
}

/** An element's child elements, as XML. */
function contentOf(element: Element | undefined) {
	return [...(element?.children ?? [])]
		.map((child) => child.outerHTML)
		.join("");
}
