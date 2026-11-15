/**
 * PowerPoint for Mac applies a font's GPOS pair kerning only some of the time:
 * a freshly launched PowerPoint often draws kerned runs unkerned, until it has
 * opened certain documents (like a full original talk), and then keeps
 * kerning until it quits. The same file then renders a few percent wider or
 * narrower from one PowerPoint process to the next, so renders are only
 * comparable once PowerPoint is known to kern.
 */
import JSZip from "jszip";
import { execFile } from "node:child_process";
import fs from "node:fs/promises";
import path from "node:path";
import { promisify } from "node:util";
import pptxgen from "pptxgenjs";

import { deckPath } from "../corpus/paths.ts";

const run = promisify(execFile);

/** A corpus talk that, opened in PowerPoint, has turned its kerning on (observed in every full render of it). */
const WARM_UP_DECK = deckPath(
	"2023-06-22-static-analysis-don-t-fear-the-linter",
);

/** How many times to open the warm-up deck before giving up on kerning. */
const WARM_UP_ATTEMPTS = 3;

/** Kerning narrows the probe's text by about 3%; less than this is measuring noise. */
const MIN_KERNED_NARROWING = 0.01;

/** A font installed for the corpus whose kerning pairs are in GPOS, with text full of them. */
const PROBE_FONT = "Regular Bold";
const PROBE_TEXT = "AVATAR To Wa Ty";

/**
 * Makes sure the running PowerPoint kerns text before a render, opening the
 * warm-up deck until it does, once per PowerPoint process.
 * @param directory PowerPoint's own container directory, for scratch files.
 * @param exportPdf Has PowerPoint save a deck as a PDF.
 * @param pid The running PowerPoint's process id, if it's running.
 */
export async function ensureKerning(
	directory: string,
	exportPdf: (input: string, pdf: string) => Promise<void>,
	pid: () => Promise<number | undefined>,
) {
	const record = path.join(directory, "kerning.json");
	const checked = JSON.parse(
		await fs.readFile(record, "utf8").catch(() => "{}"),
	) as { pid?: number };
	const before = await pid();
	if (before !== undefined && checked.pid === before) {
		return;
	}

	const probe = path.join(directory, "kerning-probe.pptx");
	await writeProbe(probe);
	const kerns = async () => {
		const pdf = path.join(directory, "kerning-probe.pdf");
		await exportPdf(probe, pdf);
		const widths = await textWidths(pdf, directory);
		await fs.rm(pdf, { force: true });
		return widths.kerned < widths.unkerned * (1 - MIN_KERNED_NARROWING);
	};

	let kerning = await kerns();
	for (
		let attempt = 0;
		!kerning &&
		attempt < WARM_UP_ATTEMPTS &&
		(await fs.stat(WARM_UP_DECK).catch(() => undefined));
		attempt++
	) {
		const input = path.join(directory, "kerning-warm-up.pptx");
		const pdf = path.join(directory, "kerning-warm-up.pdf");
		await fs.copyFile(WARM_UP_DECK, input);
		await exportPdf(input, pdf);
		await fs.rm(input, { force: true });
		await fs.rm(pdf, { force: true });
		kerning = await kerns();
	}

	if (!kerning) {
		console.error(
			"PowerPoint isn't kerning text, so its renders may differ from earlier ones.",
		);
	}

	await fs.writeFile(record, JSON.stringify({ pid: await pid() }));
}

/** The width of the probe's kerned and unkerned text, rasterized. */
async function textWidths(pdf: string, directory: string) {
	const png = path.join(directory, "kerning-probe");
	await run("pdftoppm", ["-r", "72", "-png", "-singlefile", pdf, png]);
	const widthOf = async (band: string) => {
		const { stdout } = await run("magick", [
			`${png}.png`,
			"-crop",
			band,
			"+repage",
			"-fuzz",
			"20%",
			"-trim",
			"-format",
			"%w",
			"info:",
		]);
		return Number(stdout);
	};

	try {
		return {
			kerned: await widthOf("100%x50%+0+0"),
			unkerned: await widthOf("100%x50%+0+270"),
		};
	} finally {
		await fs.rm(`${png}.png`, { force: true });
	}
}

/** A slide with the same text kerned (kern="100") above and unkerned (kern="0") below. */
async function writeProbe(file: string) {
	const presentation = new (pptxgen as unknown as typeof pptxgen.default)();
	presentation.layout = "LAYOUT_WIDE";
	const slide = presentation.addSlide();
	slide.background = { color: "000000" };
	for (const [index, kern] of [100, 0].entries()) {
		slide.addText(PROBE_TEXT, {
			color: "FFFFFF",
			fontFace: PROBE_FONT,
			fontSize: 72,
			h: 1.5,
			margin: 0,
			objectName: `kern ${String(kern)}`,
			w: 12,
			x: 0.5,
			y: 0.75 + index * 3.75,
		});
	}

	await presentation.writeFile({ fileName: file });
	const zip = await JSZip.loadAsync(await fs.readFile(file));
	const slideXml = zip.file("ppt/slides/slide1.xml");
	if (!slideXml) {
		throw new Error("The kerning probe has no slide.");
	}

	let xml = await slideXml.async("string");
	for (const kern of [100, 0]) {
		xml = xml.replace(
			new RegExp(`(name="kern ${String(kern)}".*?<a:rPr)`, "s"),
			`$1 kern="${String(kern)}"`,
		);
	}

	zip.file("ppt/slides/slide1.xml", xml);
	await fs.writeFile(file, await zip.generateAsync({ type: "nodebuffer" }));
}
