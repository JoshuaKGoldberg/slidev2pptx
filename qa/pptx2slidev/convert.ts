/**
 * Converts a .pptx into a Slidev deck, as faithfully as it reasonably can.
 * This is step 1 of round-trip QA, so it may lose things; it logs what it loses.
 * @example
 * ```shell
 * node qa/pptx2slidev/convert.ts deck.pptx out/
 * ```
 */
import JSZip from "jszip";
import fs from "node:fs/promises";
import path from "node:path";
import { PptxRenderer } from "pptx-svg";

import type { Visibility } from "./timing.ts";

import { createFontResolver } from "./fonts.ts";
import { IDENTITY } from "./geometry.ts";
import { convertShape } from "./shapes.ts";
import { readTiming } from "./timing.ts";
import { convertTransition } from "./transitions.ts";
import { descendantsNamed, parseXml } from "./xml.ts";

export interface ConversionLog {
	fonts: { missing: string[] };
	slides: {
		clicks: number;
		hidden: boolean;
		losses: string[];
		number: number;
	}[];
}

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
	const renderer = new PptxRenderer();
	await renderer.init(
		await fs.readFile(
			path.join(
				import.meta.dirname,
				"../../node_modules/pptx-svg/dist/main.wasm",
			),
		),
	);
	await renderer.loadPptx(
		bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
	);

	await fs.rm(output, { force: true, recursive: true });
	await fs.mkdir(path.join(output, "public/media"), { recursive: true });

	const fonts = createFontResolver();
	const log: ConversionLog = { fonts: { missing: [] }, slides: [] };
	const presentation = parseXml(
		(await zip.file("ppt/presentation.xml")?.async("string")) ?? "",
	);
	const slidePaths = await slideOrder(zip, presentation);
	const copied = new Map<string, string>();
	const writes: Promise<void>[] = [];
	const slides: { frontmatter: string[]; html: string; notes: string }[] = [];
	const transitions: (string | undefined)[] = [];

	for (const [index, slidePath] of slidePaths.entries()) {
		const losses: string[] = [];
		const xml = (await zip.file(slidePath)?.async("string")) ?? "";
		const slide = parseXml(xml).documentElement;
		const rels = await relationships(zip, slidePath);
		const timing = readTiming(slide);
		losses.push(
			...timing.dropped.map((effect) => `Dropped animation: ${effect}`),
		);

		const raw = new Map<string, Element>();
		for (const properties of descendantsNamed(slide, "p:cNvPr")) {
			const shape = properties.parentElement?.parentElement;
			const id = properties.getAttribute("id");
			if (shape && id) {
				raw.set(id, shape);
			}
		}

		const svg = parseXml(renderer.renderSlideSvg(index)).documentElement;
		const defs = [...svg.children]
			.filter((child) => child.tagName === "defs")
			.map((child) => child.outerHTML)
			.join("");
		const context = {
			copyMedia: (relationshipId: string) => {
				const target = rels.get(relationshipId);
				if (!target || /^https?:/.test(target)) {
					return target;
				}

				const existing = copied.get(target);
				if (existing) {
					return existing;
				}

				const name = `/media/${path.basename(target)}`;
				copied.set(target, name);
				const file = zip.file(target);
				if (!file) {
					return undefined;
				}

				writes.push(
					file
						.async("nodebuffer")
						.then((buffer) =>
							fs.writeFile(path.join(output, "public", name), buffer),
						),
				);
				return name;
			},
			defs,
			fonts,
			lose: (loss: string) => losses.push(loss),
			raw,
			resolveLink: (id: string) => rels.get(id),
			timing,
			withVisibility,
		};

		const background = backgroundOf(svg);
		const shapes = [...svg.children]
			.filter((child) => child.hasAttribute("data-ooxml-shape-type"))
			.map((child) => convertShape(context, child, IDENTITY, []))
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

	await Promise.all(writes);
	await fs.writeFile(path.join(output, "slides.md"), markdown);
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

function backgroundOf(svg: Element) {
	const rect = [...svg.children].find(
		(child) =>
			child.tagName === "rect" &&
			child.getAttribute("width") === svg.getAttribute("width"),
	);
	const fill = rect?.getAttribute("fill");
	return fill && !fill.startsWith("url(") ? `background-color:${fill}` : "";
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
			await fs.copyFile(face.file, path.join(output, "fonts", name));
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

.pptx-text {
	display: flex;
	flex-direction: column;
	overflow-wrap: break-word;
	white-space: pre-wrap;
}

.pptx-text p,
.pptx-text ul,
.pptx-text ol,
.pptx-text li {
	margin: 0;
	padding: 0;
}

.pptx-text ul,
.pptx-text ol {
	list-style-position: outside;
}

.pptx-text a {
	color: inherit;
}
`;
}

function withVisibility(html: string, visibility: undefined | Visibility) {
	if (!visibility) {
		return html;
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
