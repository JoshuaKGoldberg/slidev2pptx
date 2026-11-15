/**
 * Converts a .pptx into a Slidev deck, as faithfully as it reasonably can.
 * This is step 1 of round-trip QA, so it may lose things; it logs what it loses.
 * @example
 * ```shell
 * node qa/pptx2slidev/convert.ts deck.pptx out/
 * ```
 */
import JSZip from "jszip";
import { createHash } from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";
import { emfToSvg, PptxRenderer, wmfToSvg } from "pptx-svg";

import type { SlideContext } from "./shapes.ts";
import type { SlideTiming, Visibility } from "./timing.ts";

import { inlineDiagrams } from "./diagrams.ts";
import { createFontResolver } from "./fonts.ts";
import { IDENTITY } from "./geometry.ts";
import {
	inheritedBodyProperties,
	inheritedFrame,
	inheritedLevels,
	loadInheritance,
} from "./inheritance.ts";
import { type ImageMetrics, imageMetrics } from "./media.ts";
import { withoutEmptyTables } from "./sanitize.ts";
import { convertShape } from "./shapes.ts";
import { createTheme } from "./theme.ts";
import { readTiming } from "./timing.ts";
import { convertTransition } from "./transitions.ts";
import {
	childNamed,
	descendantNamed,
	descendantsNamed,
	parseXml,
} from "./xml.ts";

export interface ConversionLog {
	fonts: { missing: string[] };
	slides: {
		clicks: number;
		hidden: boolean;
		losses: string[];
		number: number;
	}[];
}

/** Fills a group's properties can give its children's a:grpFill. */
const GROUP_FILLS = new Set([
	"a:blipFill",
	"a:gradFill",
	"a:noFill",
	"a:pattFill",
	"a:solidFill",
]);

const [input, output] = process.argv.slice(2);
if (!input || !output) {
	throw new Error(
		"Usage: node qa/pptx2slidev/convert.ts <deck.pptx> <output directory>",
	);
}

await convertDeck(path.resolve(input), path.resolve(output));

export async function convertDeck(input: string, output: string) {
	const bytes = await fs.readFile(input);
	const zip = await JSZip.loadAsync(bytes);
	const diagrams = await inlineDiagrams(zip, (part) =>
		relationships(zip, part),
	);
	const renderer = new PptxRenderer();
	await renderer.init(
		await fs.readFile(
			path.join(
				import.meta.dirname,
				"../../node_modules/pptx-svg/dist/main.wasm",
			),
		),
	);
	await renderer.loadPptx(await forPptxSvg(zip, bytes, diagrams.size > 0));

	await fs.rm(output, { force: true, recursive: true });
	await fs.mkdir(path.join(output, "public/media"), { recursive: true });

	const fonts = createFontResolver();
	const log: ConversionLog = { fonts: { missing: [] }, slides: [] };
	const presentation = parseXml(
		(await zip.file("ppt/presentation.xml")?.async("string")) ?? "",
	);
	const slidePaths = await slideOrder(zip, presentation);
	const tableStyles = parseXml(
		(await zip.file("ppt/tableStyles.xml")?.async("string")) ??
			"<a:tblStyleLst/>",
	).documentElement;
	const themes = new Map<string, Element | undefined>();
	const media = await tiledMedia(zip);
	const copied = new Map<string, string>();
	const writes: Promise<void>[] = [];
	const slides: { frontmatter: string[]; html: string; notes: string }[] = [];
	const transitions: (string | undefined)[] = [];

	for (const [index, slidePath] of slidePaths.entries()) {
		const losses = [...(diagrams.get(slidePath) ?? [])];
		const xml = (await zip.file(slidePath)?.async("string")) ?? "";
		const slide = parseXml(xml).documentElement;
		const rels = await relationships(zip, slidePath);
		const sources = await loadInheritance(
			zip,
			slidePath,
			presentation.documentElement,
			(part) => relationships(zip, part),
		);
		const timing = readTiming(slide);
		const themePath = sources.masterPath
			? [...(await relationships(zip, sources.masterPath)).values()].find(
					(target) => target.includes("theme/"),
				)
			: undefined;
		if (themePath && !themes.has(themePath)) {
			themes.set(
				themePath,
				parseXml((await zip.file(themePath)?.async("string")) ?? "<empty/>")
					.documentElement,
			);
		}

		const theme = createTheme(
			themePath ? themes.get(themePath) : undefined,
			[slide, sources.layout, sources.master],
			tableStyles,
		);
		losses.push(
			...timing.dropped.map((effect) => `Dropped animation: ${effect}`),
		);

		const svg = parseXml(renderer.renderSlideSvg(index)).documentElement;
		// pptx-svg writes a group's children's gradients, markers, and patterns inside the group.
		const defs = new Map(
			descendantsNamed(svg, "defs")
				.flatMap((child) => [...child.children])
				.flatMap((child) => {
					const id = child.getAttribute("id");
					return id ? [[id, withStopsInOrder(child).outerHTML] as const] : [];
				}),
		);
		const copyMedia = (
			partRels: Map<string, string>,
			relationshipId: string,
		) => {
			const target = partRels.get(relationshipId);
			if (!target || /^https?:/.test(target)) {
				return target;
			}

			const existing = copied.get(target);
			if (existing) {
				return existing;
			}

			// Vite won't serve Windows metafiles, so they become SVG.
			const metafile = /\.(?:emf|wmf)$/i.exec(target)?.[0].toLowerCase();
			const name = `/media/${path.basename(target)}${metafile ? ".svg" : ""}`;
			copied.set(target, name);
			const file = zip.file(target);
			if (!file) {
				return undefined;
			}

			writes.push(
				file
					.async("nodebuffer")
					.then((buffer) =>
						fs.writeFile(
							path.join(output, "public", name),
							metafile
								? (metafile === ".wmf" ? wmfToSvg : emfToSvg)(buffer) ||
										`<svg xmlns="http://www.w3.org/2000/svg"/>`
								: buffer,
						),
					),
			);
			return name;
		};

		const contextOf = (
			partRels: Map<string, string>,
			raw: Map<string, Element>,
			partTiming: SlideTiming,
		): SlideContext => ({
			copyMedia: (relationshipId: string) =>
				copyMedia(partRels, relationshipId),
			defs,
			fonts,
			imageMetrics: (relationshipId: string) =>
				media.get(partRels.get(relationshipId) ?? ""),
			inherited: (body: Element | undefined, level: number) =>
				inheritedLevels(sources, body, level),
			inheritedBody: (body: Element | undefined) =>
				inheritedBodyProperties(sources, body),
			lose: (loss: string) => losses.push(loss),
			raw,
			resolveLink: (id: string) => partRels.get(id),
			theme,
			timing: partTiming,
			withVisibility,
		});
		const context = contextOf(rels, shapesById(slide), timing);

		// pptx-svg draws a layout's and master's own shapes before the slide's, with negative indexes.
		// Their ids can repeat the slide's, so each is read from its own part, without the slide's animations.
		const inheritedParts = await Promise.all(
			(
				[
					[sources.layout, sources.layoutPath],
					[sources.master, sources.masterPath],
				] as const
			).flatMap(([root, part]) =>
				root && part
					? [
							relationships(zip, part).then((partRels) => ({
								context: contextOf(partRels, shapesById(root), staticTiming()),
								rels: partRels,
								root,
							})),
						]
					: [],
			),
		);
		const contextFor = (shape: Element) => {
			if (Number(shape.getAttribute("data-ooxml-shape-idx") ?? 0) >= 0) {
				return context;
			}

			const id = shape.getAttribute("data-ooxml-nv-id") ?? "";
			const name = shape.getAttribute("data-ooxml-nv-name");
			const candidates = inheritedParts.filter(({ context }) =>
				context.raw.has(id),
			);
			const match =
				candidates.find(
					({ context }) =>
						descendantNamed(context.raw.get(id), "p:cNvPr")?.getAttribute(
							"name",
						) === name,
				) ?? candidates.at(0);
			if (!match) {
				losses.push(
					`Couldn't find layout or master shape ${id} ("${name ?? ""}"), so it's drawn without its text and pictures`,
				);
			}

			return match?.context ?? contextOf(rels, new Map(), staticTiming());
		};

		const background = backgroundOf(svg, defs, {
			copyMedia,
			lose: (loss) => losses.push(loss),
			parts: [{ rels, root: slide }, ...inheritedParts],
		});
		const shapes = [...svg.children]
			.filter((child) => child.hasAttribute("data-ooxml-shape-type"))
			.map((child) => convertShape(contextFor(child), child, IDENTITY, []))
			.join("\n");

		const hidden = renderer.isSlideHidden(index);
		const frontmatter = ["layout: none"];
		if (hidden) {
			frontmatter.push("hide: true");
		}

		const transition = convertTransition(slide);
		if (transition?.loss) {
			losses.push(transition.loss);
		}

		transitions.push(transition?.value);
		slides.push({
			frontmatter,
			// A blank line would end Markdown's HTML block, so the slide's HTML never has one.
			html: `<div class="pptx" style="${background}">\n${shapes}\n</div>`.replaceAll(
				/\n\s*\n/g,
				"\n",
			),
			notes: renderer.getSlideNotes(index).join("\n").trim(),
		});
		log.slides.push({
			clicks: timing.clicks,
			hidden,
			losses,
			number: index + 1,
		});
	}

	// PowerPoint transitions into a slide; Slidev transitions out of one.
	for (const [index, transition] of transitions.entries()) {
		if (index > 0) {
			slides[index - 1].frontmatter.push(`transition: ${transition ?? "none"}`);
		}
	}

	const title =
		/<dc:title>([^<]*)<\/dc:title>/.exec(
			(await zip.file("docProps/core.xml")?.async("string")) ?? "",
		)?.[1] ?? path.basename(input, ".pptx");
	const markdown = slides
		.map((slide, index) => {
			const frontmatter =
				index === 0
					? [
							"theme: none",
							`title: ${JSON.stringify(title)}`,
							"aspectRatio: 16/9",
							"canvasWidth: 960",
							"colorSchema: light",
							"fonts:",
							"  provider: none",
							...slide.frontmatter,
						]
					: slide.frontmatter;
			const notes = slide.notes
				? `\n\n<!--\n${slide.notes.replaceAll("-->", "--&gt;")}\n-->`
				: "";
			return `---\n${frontmatter.join("\n")}\n---\n\n${slide.html}${notes}\n`;
		})
		.join("\n");

	// pptx-svg embeds picture fills as data URIs, which can make Markdown too big for Vite to compile.
	const embedded = new Set<string>();
	const external = markdown.replaceAll(
		/data:image\/(gif|jpeg|png|svg\+xml|webp);base64,([A-Za-z0-9+/=]+)/g,
		(_, type: string, data: string) => {
			const hash = createHash("sha256").update(data).digest("hex").slice(0, 16);
			const name = `/media/embedded-${hash}.${type.replace(/\+xml$/, "")}`;
			if (!embedded.has(name)) {
				embedded.add(name);
				writes.push(
					fs.writeFile(
						path.join(output, "public", name),
						Buffer.from(data, "base64"),
					),
				);
			}

			return name;
		},
	);

	await Promise.all(writes);
	await fs.writeFile(path.join(output, "slides.md"), external);
	await fs.writeFile(
		path.join(output, "style.css"),
		await styleSheet(fonts.faces(), output),
	);
	log.fonts.missing = fonts.missing();
	await fs.writeFile(
		path.join(output, "conversion-log.json"),
		JSON.stringify(log, null, "\t") + "\n",
	);

	console.log(
		`Converted ${String(slides.length)} slides to ${output}; ${String(log.slides.reduce((sum, slide) => sum + slide.losses.length, 0))} losses logged.`,
	);
}

/** Where a slide's background can come from, most specific first: the slide, its layout, then its master. */
interface BackgroundPart {
	rels: Map<string, string>;
	root: Element;
}

/** The slide's background as inline CSS for its root element. */
function backgroundOf(
	svg: Element,
	defs: Map<string, string>,
	{
		copyMedia,
		lose,
		parts,
	}: {
		copyMedia: (
			rels: Map<string, string>,
			relationshipId: string,
		) => string | undefined;
		lose: (loss: string) => void;
		parts: BackgroundPart[];
	},
) {
	const width = Number(svg.getAttribute("width"));
	const height = Number(svg.getAttribute("height"));

	// pptx-svg leaves picture backgrounds undrawn, so they're read from whichever part sets the background.
	const owner = parts.find(({ root }) =>
		childNamed(childNamed(root, "p:cSld"), "p:bg"),
	);
	const blipFill = descendantNamed(
		childNamed(childNamed(owner?.root, "p:cSld"), "p:bg"),
		"a:blipFill",
	);
	if (owner && blipFill) {
		return pictureBackground(
			blipFill,
			{ height, width },
			(id) => copyMedia(owner.rels, id),
			lose,
		);
	}

	const rect = [...svg.children].find(
		(child) =>
			child.tagName === "rect" &&
			child.getAttribute("width") === svg.getAttribute("width"),
	);
	const fill = rect?.getAttribute("fill");
	if (!fill || fill === "none") {
		return "";
	}

	const reference = /^url\(#([^)]+)\)$/.exec(fill)?.[1];
	if (!reference) {
		return `background-color:${fill}`;
	}

	const markup = defs.get(reference);
	const gradient = markup ? cssGradient(markup, { height, width }) : undefined;
	if (!gradient) {
		lose(`Dropped the slide's background fill ${reference}`);
		return "";
	}

	return `background-image:${gradient}`;
}

/** A background picture, stretched into its fill rectangle and cropped by its source rectangle, as CSS. */
function pictureBackground(
	blipFill: Element,
	slide: { height: number; width: number },
	copyMedia: (relationshipId: string) => string | undefined,
	lose: (loss: string) => void,
) {
	const blip = childNamed(blipFill, "a:blip");
	const url = copyMedia(blip?.getAttribute("r:embed") ?? "");
	if (!url) {
		lose("Dropped the slide's background picture, as it couldn't be copied");
		return "";
	}

	if (childNamed(blipFill, "a:tile")) {
		lose("Stretched the slide's tiled background picture");
	}

	if (
		blip &&
		[...blip.children].some(
			(effect) => effect.tagName !== "a:lum" || effect.attributes.length,
		)
	) {
		lose(
			`Dropped ${[...blip.children].map((effect) => effect.tagName.replace("a:", "")).join(", ")} on the slide's background picture`,
		);
	}

	// Both rectangles' edges are insets in thousandths of a percent; negative ones reach past the slide.
	const edges = (element: Element | undefined) =>
		["l", "t", "r", "b"].map(
			(edge) => Number(element?.getAttribute(edge) ?? 0) / 100_000,
		);
	const [left, top, right, bottom] = edges(
		childNamed(childNamed(blipFill, "a:stretch"), "a:fillRect"),
	);
	const [cropLeft, cropTop, cropRight, cropBottom] = edges(
		childNamed(blipFill, "a:srcRect"),
	);
	const shownWidth = slide.width * (1 - left - right);
	const shownHeight = slide.height * (1 - top - bottom);
	const imageWidth = shownWidth / (1 - cropLeft - cropRight);
	const imageHeight = shownHeight / (1 - cropTop - cropBottom);
	const round = (value: number) => `${String(Math.round(value * 100) / 100)}px`;
	return [
		`background-image:url(${url})`,
		`background-position:${round(slide.width * left - cropLeft * imageWidth)} ${round(slide.height * top - cropTop * imageHeight)}`,
		"background-repeat:no-repeat",
		`background-size:${round(imageWidth)} ${round(imageHeight)}`,
	].join(";");
}

/**
 * A pptx-svg linear gradient over the slide as a CSS linear-gradient().
 * pptx-svg rotates gradients in their bounding box, which stretches to the slide,
 * so the CSS angle is the bounding box direction's normal in slide pixels.
 */
function cssGradient(markup: string, slide: { height: number; width: number }) {
	const gradient = parseXml(markup).documentElement;
	if (
		gradient.tagName !== "linearGradient" ||
		gradient.getAttribute("gradientUnits") === "userSpaceOnUse"
	) {
		return undefined;
	}

	const coordinate = (name: string, fallback: number) => {
		const value = gradient.getAttribute(name);
		if (!value) {
			return fallback;
		}

		return value.endsWith("%") ? Number.parseFloat(value) / 100 : Number(value);
	};

	const turn =
		(Number(
			/rotate\(\s*(-?[\d.]+)/.exec(
				gradient.getAttribute("gradientTransform") ?? "",
			)?.[1] ?? 0,
		) *
			Math.PI) /
		180;
	const run = coordinate("x2", 1) - coordinate("x1", 0);
	const rise = coordinate("y2", 0) - coordinate("y1", 0);
	const dx = run * Math.cos(turn) - rise * Math.sin(turn);
	const dy = run * Math.sin(turn) + rise * Math.cos(turn);
	const degrees =
		((Math.atan2(dx * slide.height, -dy * slide.width) * 180) / Math.PI + 360) %
		360;

	// CSS clamps a stop placed before an earlier one, so stops go in order; pptx-svg keeps OOXML's, which needn't be.
	const stops = [...gradient.children]
		.filter((child) => child.tagName === "stop")
		.map((stop) => {
			const offset = stop.getAttribute("offset") ?? "0";
			const position = offset.endsWith("%")
				? Number.parseFloat(offset)
				: Number(offset) * 100;
			const color = stop.getAttribute("stop-color") ?? "#000000";
			const opacity = Number(stop.getAttribute("stop-opacity") ?? 1);
			return { color: withOpacity(color, opacity), position };
		})
		.sort((a, b) => a.position - b.position)
		.map(
			({ color, position }) =>
				`${color} ${String(Math.round(position * 1000) / 1000)}%`,
		);
	if (!stops.length) {
		return undefined;
	}

	return `linear-gradient(${String(Math.round(degrees * 1000) / 1000)}deg, ${stops.join(", ")})`;
}

/**
 * A gradient with its stops in order of offset, as PowerPoint takes them, where
 * pptx-svg keeps OOXML's order and SVG moves a stop placed before an earlier one up to it.
 */
function withStopsInOrder(gradient: Element) {
	const offset = (stop: Element) => {
		const value = stop.getAttribute("offset") ?? "0";
		return value.endsWith("%") ? Number.parseFloat(value) / 100 : Number(value);
	};

	const stops = [...gradient.children].filter(
		(child) => child.tagName === "stop",
	);
	for (const stop of stops.toSorted((a, b) => offset(a) - offset(b))) {
		gradient.append(stop);
	}

	return gradient;
}

/** A CSS color faded by an opacity, as rgba() when it's a hex color. */
function withOpacity(color: string, opacity: number) {
	const hex = /^#([0-9a-f]{6})$/i.exec(color)?.[1];
	if (opacity >= 1 || !hex) {
		return color;
	}

	const channels = [0, 2, 4].map((start) =>
		Number.parseInt(hex.slice(start, start + 2), 16),
	);
	return `rgba(${channels.join(", ")}, ${String(opacity)})`;
}

/** Pixel sizes and resolutions of a deck's images, which only its tiled picture fills need, so only decks with one read them. */
async function tiledMedia(zip: JSZip) {
	const media = new Map<string, ImageMetrics>();
	const parts = Object.keys(zip.files).filter((name) =>
		/^ppt\/(?:slides|slideLayouts|slideMasters)\/[^/]+\.xml$/.test(name),
	);
	const tiled = await Promise.all(
		parts.map(async (name) =>
			((await zip.file(name)?.async("string")) ?? "").includes("<a:tile"),
		),
	);
	if (!tiled.some(Boolean)) {
		return media;
	}

	for (const name of Object.keys(zip.files)) {
		const file = zip.file(name);
		if (name.startsWith("ppt/media/") && file) {
			const metrics = imageMetrics(await file.async("nodebuffer"));
			if (metrics) {
				media.set(name, metrics);
			}
		}
	}

	return media;
}

/** A part's shape elements by their cNvPr id, including shapes in groups. */
function shapesById(root: Element) {
	const shapes = new Map<string, Element>();
	for (const properties of descendantsNamed(root, "p:cNvPr")) {
		const shape = properties.parentElement?.parentElement;
		const id = properties.getAttribute("id");
		if (shape && id) {
			shapes.set(id, shape);
		}
	}

	return shapes;
}

/** Timing for layout and master shapes, which slide animations never target. */
function staticTiming(): SlideTiming {
	return {
		clicks: 0,
		dropped: [],
		effects: new Map(),
		paragraphs: new Map(),
		shapes: new Map(),
	};
}

/**
 * The deck with what pptx-svg mishandles rewritten:
 * - Characters outside the BMP in presentation.xml, such as an emoji in a section name, are escaped.
 *   pptx-svg 0.6.8 crashes loading a deck with one, but not when it's written as a reference.
 * - Group fills (a:grpFill), which pptx-svg leaves empty, become their group's fill.
 * @see https://github.com/t-ujiie-g/pptx-svg/issues/88
 */
async function forPptxSvg(
	zip: JSZip,
	bytes: Buffer<ArrayBuffer>,
	/** Whether the deck's parts were already changed, as when its diagrams are inlined. */
	changed = false,
) {
	// Placeholder pictures without frames take the frames they inherit, as pptx-svg skips them.
	const framed = await Promise.all(
		Object.keys(zip.files)
			.filter((name) => /^ppt\/slides\/[^/]+\.xml$/.test(name))
			.map((name) => framePlaceholderPictures(zip, name)),
	);
	const rewrites = new Map<string, string>();
	const part = "ppt/presentation.xml";
	const xml = (await zip.file(part)?.async("string")) ?? "";
	const escaped = xml.replaceAll(
		/[\u{10000}-\u{10FFFF}]/gu,
		(character) => `&#x${(character.codePointAt(0) ?? 0).toString(16)};`,
	);
	if (escaped !== xml) {
		rewrites.set(part, escaped);
	}

	for (const name of Object.keys(zip.files).filter((name) =>
		/^ppt\/(?:slides|slideLayouts|slideMasters)\/[^/]+\.xml$/.test(name),
	)) {
		const original = (await zip.file(name)?.async("string")) ?? "";
		let rewritten = withGroupFills(original);
		// PowerPoint matches a layout's placeholder to its master's by type, where pptx-svg matches by index
		// first, so a layout's second content placeholder (idx 2) would take the master's date's 12pt text.
		if (name.startsWith("ppt/slideMasters/")) {
			rewritten = rewritten.replaceAll(/<p:ph\b[^>]*>/g, (tag) =>
				tag.replace(/ idx="\d+"/, ""),
			);
		}

		if (rewritten !== original) {
			rewrites.set(name, rewritten);
		}
	}

	if (!changed && !rewrites.size && !framed.includes(true)) {
		return bytes.buffer.slice(
			bytes.byteOffset,
			bytes.byteOffset + bytes.byteLength,
		);
	}

	for (const [name, rewritten] of rewrites) {
		zip.file(name, rewritten);
	}

	return await zip.generateAsync({ type: "arraybuffer" });
}

/** Writes the frames a slide's placeholder pictures inherit into those without one; returns whether it changed the slide. */
async function framePlaceholderPictures(zip: JSZip, part: string) {
	const xml = (await zip.file(part)?.async("string")) ?? "";
	if (!xml.includes("<p:pic") || !xml.includes("<p:ph")) {
		return false;
	}

	const pictures = descendantsNamed(
		parseXml(xml).documentElement,
		"p:pic",
	).flatMap((picture) => {
		const id = descendantNamed(picture, "p:cNvPr")?.getAttribute("id");
		const placeholder = descendantNamed(
			childNamed(picture, "p:nvPicPr"),
			"p:ph",
		);
		return id &&
			placeholder &&
			!childNamed(childNamed(picture, "p:spPr"), "a:xfrm")
			? [{ id, placeholder }]
			: [];
	});
	if (!pictures.length) {
		return false;
	}

	const sources = await loadInheritance(zip, part, undefined, (name) =>
		relationships(zip, name),
	);
	let framed = xml;
	// Alternate content can repeat a picture, but only its first copy is found by its id.
	for (const [id, placeholder] of new Map(
		pictures.map(({ id, placeholder }) => [id, placeholder]),
	)) {
		const frame = inheritedFrame(sources, placeholder);
		// The picture's shape properties are the first after its id, and its frame goes first in them.
		const at = framed.search(new RegExp(`<p:cNvPr\\b[^>]*\\bid="${id}"`));
		const opening = /<p:spPr\b([^>]*?)(\/?)>/g;
		opening.lastIndex = at;
		const match = at >= 0 ? opening.exec(framed) : null;
		if (frame && match) {
			framed = `${framed.slice(0, match.index)}<p:spPr${match[1]}>${frame.outerHTML}${match[2] ? "</p:spPr>" : ""}${framed.slice(opening.lastIndex)}`;
		}
	}

	if (framed === xml) {
		return false;
	}

	zip.file(part, framed);
	return true;
}

async function relationships(zip: JSZip, part: string) {
	const relsPath = path.posix.join(
		path.posix.dirname(part),
		"_rels",
		`${path.posix.basename(part)}.rels`,
	);
	const xml = parseXml(
		(await zip.file(relsPath)?.async("string")) ?? "<Relationships/>",
	);
	const targets = new Map<string, string>();
	for (const relationship of descendantsNamed(
		xml.documentElement,
		"Relationship",
	)) {
		const id = relationship.getAttribute("Id");
		const target = relationship.getAttribute("Target");
		if (id && target) {
			targets.set(
				id,
				relationship.getAttribute("TargetMode") === "External"
					? target
					: path.posix.normalize(
							path.posix.join(path.posix.dirname(part), target),
						),
			);
		}
	}

	return targets;
}

async function slideOrder(zip: JSZip, presentation: Document) {
	const rels = await relationships(zip, "ppt/presentation.xml");
	return descendantsNamed(presentation.documentElement, "p:sldId").flatMap(
		(slide) => {
			const target = rels.get(slide.getAttribute("r:id") ?? "");
			return target ? [target] : [];
		},
	);
}

async function styleSheet(
	faces: {
		family: string;
		file: string;
		italic: boolean;
		weight: [number, number] | number;
	}[],
	output: string,
) {
	await fs.mkdir(path.join(output, "fonts"), { recursive: true });
	const rules: string[] = [];
	const files = new Map<string, string>();
	for (const face of faces) {
		let name = files.get(face.file);
		if (!name) {
			name = `${String(files.size)}-${path.basename(face.file).replaceAll(/[^\w.-]/g, "_")}`;
			files.set(face.file, name);
			await fs.writeFile(
				path.join(output, "fonts", name),
				withoutEmptyTables(await fs.readFile(face.file)),
			);
		}

		rules.push(
			`@font-face {\n\tfont-family: "${face.family}";\n\tsrc: url("./fonts/${name}");\n\tfont-weight: ${Array.isArray(face.weight) ? face.weight.join(" ") : String(face.weight)};\n\tfont-style: ${face.italic ? "italic" : "normal"};\n}`,
		);
	}

	return `${[...new Set(rules)].join("\n\n")}

.pptx {
	position: absolute;
	inset: 0;
	overflow: hidden;
	font-family: Calibri, sans-serif;
	color: black;
	text-align: left;
}

.pptx-group {
	position: absolute;
	inset: 0;
	pointer-events: none;
}

.pptx-shape {
	position: absolute;
	box-sizing: border-box;
	margin: 0;
}

.pptx-crop {
	overflow: hidden;
}

/* UnoCSS's preflight shrinks images to their container, which would undo crops. */
.pptx img {
	max-width: none;
	max-height: none;
}

.pptx-text {
	display: flex;
	flex-direction: column;
	overflow-wrap: break-word;
	white-space: pre-wrap;
}

/* PowerPoint draws a table's borders centered on its grid, as collapsed borders are. */
.pptx-table {
	border-collapse: collapse;
	border-spacing: 0;
	table-layout: fixed;
}

.pptx-table td {
	overflow-wrap: break-word;
	white-space: pre-wrap;
}

.pptx-text p,
.pptx-text ul,
.pptx-text ol,
.pptx-text li,
.pptx-table p,
.pptx-table ul,
.pptx-table ol,
.pptx-table li {
	margin: 0;
	padding: 0;
}

.pptx-text ul,
.pptx-text ol,
.pptx-table ul,
.pptx-table ol {
	list-style-position: outside;
}

.pptx-text a,
.pptx-table a {
	color: inherit;
}
`;
}

/**
 * A part's XML with each a:grpFill replaced by its nearest group's own fill, or no fill without one.
 * PowerPoint spans a group's gradient over the whole group, where this spans it over each shape.
 */
function withGroupFills(xml: string) {
	if (!xml.includes("<a:grpFill")) {
		return xml;
	}

	const fills = descendantsNamed(
		parseXml(xml).documentElement,
		"a:grpFill",
	).map((element) => {
		// Start above the shape (or group) whose properties hold this fill; a group without a fill has none.
		let group = element.parentElement?.parentElement?.parentElement;
		for (; group; group = group.parentElement) {
			const properties = childNamed(group, "p:grpSpPr");
			const fill = [...(properties?.children ?? [])].find(
				(child) =>
					GROUP_FILLS.has(child.tagName) || child.tagName === "a:grpFill",
			);
			if (properties && fill?.tagName !== "a:grpFill") {
				return fill?.outerHTML ?? "<a:noFill/>";
			}
		}

		return "<a:noFill/>";
	});
	let index = 0;
	return xml.replaceAll(
		/<a:grpFill\s*\/>|<a:grpFill\s*>\s*<\/a:grpFill>/g,
		() => fills[index++] ?? "<a:noFill/>",
	);
}

function withVisibility(html: string, visibility: undefined | Visibility) {
	if (!visibility) {
		return html;
	}

	// An exit as the slide starts means the shape never shows.
	if (visibility.exit === 0) {
		return "";
	}

	let wrapped = html;
	if (visibility.exit !== undefined) {
		wrapped = `<div v-click.hide="${String(visibility.exit)}" class="pptx-group">${wrapped}</div>`;
	}

	if (visibility.enter) {
		wrapped = `<div v-click="${String(visibility.enter)}" class="pptx-group">${wrapped}</div>`;
	}

	return wrapped;
}
