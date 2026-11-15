import JSZip from "jszip";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { afterAll, beforeAll, describe, expect, it } from "vitest";

import type { FontBook } from "./fonts.ts";
import type {
	Deck,
	ImageItem,
	Item,
	ListItem,
	ListParagraph,
	Run,
	Slide,
	TextItem,
} from "./model.ts";

import {
	codePitch,
	dashPattern,
	horizontalPlacement,
	markerFor,
	numberLength,
	timeline,
	writePptx,
} from "./write.ts";

const fonts: FontBook = {
	// Only the font named "Measured" is loaded, with every character 0.5em wide.
	advance: (font, text) =>
		font === "Measured" ? Array.from(text).length * 0.5 : undefined,
	load: () => Promise.resolve([]),
	metrics: () => ({
		ascent: 0.8,
		baseline: 0.75,
		descent: 0.2,
		singleDescent: 0.2,
	}),
	missingFamilies: () => Promise.resolve([]),
	resolve: () => ({ bold: false, font: "Body" }),
	spacer: () => ({ advance: 0.6, font: "Mono" }),
};

const photo = fileURLToPath(
	new URL("fixture/public/photo.png", import.meta.url),
);

const style = {
	align: "left",
	available: { left: 0, right: 920 },
	lineHeight: 36,
	rtl: false,
} as const;

const run = (text: string, overrides: Partial<Run> = {}): Run => ({
	bold: false,
	color: "DAE7F3",
	font: "Body",
	italic: false,
	letterSpacing: 0,
	pad: false,
	size: 24,
	strike: false,
	text,
	transparency: 0,
	upper: false,
	...overrides,
});

const textItem = (
	paragraphs: Run[][],
	overrides: Partial<TextItem> = {},
): TextItem => ({
	baselines: [40],
	box: { height: 40, width: 400, x: 40, y: 20 },
	decorative: false,
	heading: false,
	key: "text",
	kind: "text",
	paragraphs: paragraphs.map((runs) => ({ runs })),
	rotation: 0,
	style: { ...style, nowrap: false, wrapWidth: 400 },
	...overrides,
});

const listParagraph = (
	text: string,
	overrides: Partial<ListParagraph> = {},
): ListParagraph => ({
	baselines: [],
	bullet: true,
	indent: 0,
	level: 0,
	marker: "disc",
	number: 1,
	runs: [run(text)],
	visible: true,
	...overrides,
});

const listItem = (
	paragraphs: ListParagraph[],
	overrides: Partial<ListItem> = {},
): ListItem => ({
	box: { height: 80, width: 400, x: 40, y: 20 },
	decorative: false,
	key: "list",
	kind: "list",
	paragraphs,
	style: { ...style, nowrap: false },
	...overrides,
});

const firstLevel = (xml: string | undefined, style: string) => {
	const match = new RegExp(
		`<${style}>.*?<a:lvl1pPr\\b[^>]*algn="(\\w+)"[^>]*>(.*?)</a:lvl1pPr>`,
		"s",
	).exec(xml ?? "");
	return match ? { align: match[1], properties: match[2] } : undefined;
};

const numberedList = (numbers: number[], marker: string) =>
	listItem(
		numbers.map((number) =>
			listParagraph(`Item ${String(number)}`, { marker, number }),
		),
	);

type BoxItem = Extract<Item, { kind: "box" }>;

const boxItem = (overrides: Partial<BoxItem> = {}): BoxItem => ({
	box: { height: 10, width: 10, x: 0, y: 0 },
	decorative: false,
	key: "box",
	kind: "box",
	radius: 0,
	rotation: 0,
	...overrides,
});

const imageItem = (overrides: Partial<ImageItem> = {}): ImageItem => ({
	alt: "A photo",
	box: { height: 100, width: 200, x: 0, y: 0 },
	decorative: false,
	file: photo,
	key: "image",
	kind: "image",
	radius: 0,
	src: "/photo.png",
	transparency: 0,
	...overrides,
});

describe(writePptx, () => {
	let directory: string;
	let background: string;
	let written = 0;

	beforeAll(async () => {
		directory = await fs.mkdtemp(path.join(os.tmpdir(), "pptx-write-"));
		background = path.join(directory, "background.png");
		await fs.writeFile(
			background,
			Buffer.concat([await fs.readFile(photo), Buffer.from([0])]),
		);
	});

	afterAll(async () => {
		await fs.rm(directory, { force: true, recursive: true });
	});

	const deckOf = (slides: Partial<Slide>[]): Deck => ({
		aspectRatio: 16 / 9,
		author: "Author",
		background: { color: "101010", file: background },
		canvas: { height: 540, width: 960 },
		slides: slides.map((slide, index) => ({
			notes: "",
			number: index + 1,
			states: [],
			...slide,
		})),
		title: "Deck Title",
	});

	const write = async (deck: Deck, withClicks = false) => {
		const output = path.join(directory, `${String(written++)}.pptx`);
		const result = await writePptx({ deck, fonts, output, withClicks });
		return { file: await fs.readFile(output), result };
	};

	const writeZip = async (deck: Deck, withClicks = false) =>
		JSZip.loadAsync((await write(deck, withClicks)).file);

	const writeItems = (items: Item[]) =>
		writeZip(deckOf([{ states: [{ clicks: 0, items }] }]));

	const part = (zip: JSZip, name: string) => zip.file(name)?.async("string");

	const slideXml = async (items: Item[]) =>
		part(await writeItems(items), "ppt/slides/slide1.xml");

	const notesXml = async (notes: string) =>
		part(
			await writeZip(deckOf([{ notes, states: [{ clicks: 0, items: [] }] }])),
			"ppt/notesSlides/notesSlide1.xml",
		);

	it("writes the same bytes when written twice", async () => {
		const deck = deckOf([
			{ states: [{ clicks: 0, items: [textItem([[run("Same")]])] }] },
		]);

		expect((await write(deck)).file.equals((await write(deck)).file)).toBe(
			true,
		);
	});

	it("strips control characters when text has them", async () => {
		const xml = await slideXml([textItem([[run("a\u0007b")]])]);

		expect({
			name: xml?.match(/name="Text: [^"]*"/)?.[0],
			text: xml?.match(/<a:t>[^<]*<\/a:t>/)?.[0],
		}).toEqual({ name: 'name="Text: ab"', text: "<a:t>ab</a:t>" });
	});

	it("cuts long names at a word with an ellipsis", async () => {
		expect(
			await slideXml([
				textItem([
					[
						run(
							"Every word of this text makes the name much longer than it could be",
						),
					],
				]),
			]),
		).toContain('name="Text: Every word of this text makes the name…"');
	});

	it("cuts a long name inside its first word when the word is too long to fit", async () => {
		expect(
			await slideXml([
				textItem([
					[run("one-two-three-four-five-six-seventy-twelve-thirteen")],
				]),
			]),
		).toContain('name="Text: one-two-three-four-five-six-seventy-twelve…"');
	});

	it("numbers names when shapes would share one", async () => {
		expect(
			await slideXml([
				textItem([[run("Twice")]], { key: "a" }),
				textItem([[run("Twice")]], { key: "b" }),
			]),
		).toContain('name="Text: Twice (2)"');
	});

	it("styles a blank paragraph like the text when it has no runs", async () => {
		expect(
			await slideXml([
				{
					baselines: [40, 76, 112],
					box: { height: 108, width: 400, x: 40, y: 20 },
					decorative: false,
					key: "code",
					kind: "code",
					paragraphs: [
						{ runs: [run("one", { color: "AB5959", font: "Mono" })] },
						{ runs: [] },
						{ runs: [run("two", { font: "Mono" })] },
					],
					style: { ...style, nowrap: true },
				},
			]),
		).toContain(
			'<a:p><a:pPr algn="l" indent="0" marL="0"><a:lnSpc><a:spcPct val="125000"/></a:lnSpc><a:buNone/></a:pPr><a:r><a:rPr lang="en-US" kern="100" sz="2400" dirty="0"><a:solidFill><a:srgbClr val="AB5959"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:rPr><a:t> </a:t></a:r><a:endParaRPr lang="en-US" kern="100" sz="2400" dirty="0"><a:solidFill><a:srgbClr val="AB5959"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:endParaRPr></a:p>',
		);
	});

	it("spaces lines exactly when a paragraph mixes font sizes", async () => {
		const xml = await slideXml([
			textItem([[run("“", { size: 36 }), run("Quoted text")]]),
		]);

		expect(xml).toContain('<a:lnSpc><a:spcPts val="3600"/></a:lnSpc>');
		expect(xml).not.toContain("<a:spcPct");
	});

	it.each([
		["font sizes", [run("“", { size: 36 }), run("Quoted text")]],
		["fonts", [run("Run "), run("code", { font: "Mono" })]],
	])(
		"spaces lines by their measured pitch when a paragraph mixes %s",
		async (_, runs) => {
			expect(
				await slideXml([textItem([runs], { baselines: [40, 84, 128] })]),
			).toContain('<a:lnSpc><a:spcPts val="4400"/></a:lnSpc>');
		},
	);

	it("spaces lines as a multiple when a paragraph's fonts differ only in a spacer", async () => {
		expect(
			await slideXml([
				textItem([[run("   ", { font: "Mono", pad: true }), run("Indented")]]),
			]),
		).toContain('<a:lnSpc><a:spcPct val="125000"/></a:lnSpc>');
	});

	it("spaces lines as a multiple when a paragraph's sizes differ only in a superscript", async () => {
		expect(
			await slideXml([
				textItem([[run("x"), run("2", { script: "super", size: 16 })]]),
			]),
		).toContain('<a:lnSpc><a:spcPct val="125000"/></a:lnSpc>');
	});

	it.each([
		// 0.12em at 24px.
		["❌ 💖", 2.88],
		["👩‍💻", 2.88],
		["Text ❌", 0],
		["#", 0],
	])(
		"lowers %j by %s px as PowerPoint draws emoji higher",
		async (text, drop) => {
			const offsetOf = async (content: string) =>
				Number(
					[
						...((await slideXml([textItem([[run(content)]])])) ?? "").matchAll(
							/<a:off x="\d+" y="(\d+)"/g,
						),
					].at(-1)?.[1],
				);

			expect(
				((await offsetOf(text)) - (await offsetOf("Text"))) / 12700,
			).toBeCloseTo(drop, 1);
		},
	);

	it.each([
		// 0.12em, only PowerPoint's rise, at 24px.
		[90, { x: -2.88, y: 0 }],
		[180, { x: 0, y: -2.88 }],
	])(
		"lowers emoji along a text box turned %s degrees",
		async (rotation, drop) => {
			const offsetOf = async (content: string) => {
				const [, x, y] =
					[
						...(
							(await slideXml([
								textItem([[run(content)]], { baselines: undefined, rotation }),
							])) ?? ""
						).matchAll(/<a:off x="(\d+)" y="(\d+)"/g),
					].at(-1) ?? [];
				return { x: Number(x), y: Number(y) };
			};

			const emoji = await offsetOf("❌");
			const text = await offsetOf("Text");

			expect((emoji.x - text.x) / 12700).toBeCloseTo(drop.x, 1);
			expect((emoji.y - text.y) / 12700).toBeCloseTo(drop.y, 1);
		},
	);

	it.each([
		// Single spacing keeps the font's 0.2 descent share of 28.8px below the baseline.
		[28.8, 40 - (28.8 - 0.2 * 28.8)],
		// Tighter spacing takes room from above the baseline.
		[25.92, 40 - (25.92 - 0.2 * 28.8)],
		// Looser spacing puts the baseline 0.746 down the pitch.
		[43.2, 40 - 0.746 * 43.2],
		// Even just looser than single spacing.
		[30, 40 - 0.746 * 30],
	])(
		"puts the first baseline where PowerPoint does at a %s px pitch",
		async (lineHeight, top) => {
			const xml = await slideXml([
				textItem([[run("Text")]], {
					style: { ...style, lineHeight, nowrap: false, wrapWidth: 400 },
				}),
			]);
			const y = Number(
				[...(xml ?? "").matchAll(/<a:off x="\d+" y="(\d+)"/g)].at(-1)?.[1],
			);

			expect(y / 12700).toBeCloseTo(top, 1);
		},
	);

	it.each([
		// Up to single spacing for the line's largest font (72px), the 0.2 descent share of it stays below the baseline.
		[
			66.8,
			[run("Small ", { size: 54 }), run("large", { size: 60 })],
			100 - (66.8 - 0.2 * 72),
		],
		// The largest font counts, not the first.
		[
			66.8,
			[run("Large ", { size: 60 }), run("small", { size: 54 })],
			100 - (66.8 - 0.2 * 72),
		],
		// Only the first line's fonts count.
		[
			66.8,
			[
				run("Small ", { size: 54 }),
				run("large", { lineBreak: "forced", size: 60 }),
			],
			100 - 0.75 * 66.8,
		],
		// Any looser, the baseline goes the font's measured fraction down the pitch.
		[
			74,
			[run("Small ", { size: 54 }), run("large", { size: 60 })],
			100 - 0.75 * 74,
		],
	])(
		"puts an exactly spaced first baseline where PowerPoint does at a %s px pitch",
		async (lineHeight, runs, top) => {
			const xml = await slideXml([
				textItem([runs], {
					baselines: [100],
					style: { ...style, lineHeight, nowrap: false, wrapWidth: 400 },
				}),
			]);
			const y = Number(
				[...(xml ?? "").matchAll(/<a:off x="\d+" y="(\d+)"/g)].at(-1)?.[1],
			);

			expect(xml).toContain("<a:spcPts");
			expect(y / 12700).toBeCloseTo(top, 1);
		},
	);

	it.each([
		// PowerPoint wraps a box that can wrap itself, needing the space to break at the same word.
		[false, "<a:t>GitHub templates</a:t>"],
		// A box that can't wrap gets the browser's wrap as a written break, which aligns any space before it with its line.
		[true, "<a:t>GitHub</a:t></a:r><a:br/>"],
	])(
		"keeps the space a centered line wraps at only where PowerPoint wraps it (nowrap: %s)",
		async (nowrap, text) => {
			expect(
				await slideXml([
					textItem(
						[[run("GitHub "), run("templates", { lineBreak: "wrap" })]],
						{
							style: { ...style, align: "center", nowrap, wrapWidth: 400 },
						},
					),
				]),
			).toContain(text);
		},
	);

	it("puts a list's first baseline where a text box's goes", async () => {
		const offsetOf = async (item: Item) =>
			[
				...((await slideXml([item])) ?? "").matchAll(
					/<a:off x="\d+" y="(\d+)"/g,
				),
			].at(-1)?.[1];

		expect(
			await offsetOf(listItem([listParagraph("Text", { baselines: [40] })])),
		).toBe(await offsetOf(textItem([[run("Text")]])));
	});

	it("throws when a text item has no runs", async () => {
		await expect(slideXml([textItem([[]])])).rejects.toThrow(
			"Text items need at least one run.",
		);
	});

	it("outlines an ellipse when it has a border", async () => {
		expect(
			await slideXml([
				{
					border: { color: "FF6347", transparency: 0, width: 4 },
					box: { height: 40, width: 40, x: 10, y: 10 },
					decorative: false,
					key: "ring",
					kind: "ellipse",
				},
			]),
		).toContain(
			'<a:prstGeom prst="ellipse"><a:avLst></a:avLst></a:prstGeom><a:ln w="50800"><a:solidFill><a:srgbClr val="FF6347"/></a:solidFill>',
		);
	});

	it.each([
		// CSS draws a border inside its box.
		[
			"box",
			boxItem({ box: { height: 40, width: 60, x: 10, y: 10 }, radius: 8 }),
			[12, 12, 56, 36],
		],
		// SVG centers a stroke on its shape's edge, as PowerPoint does.
		[
			"ellipse",
			{
				box: { height: 40, width: 60, x: 10, y: 10 },
				decorative: false,
				key: "ring",
				kind: "ellipse",
			} as const,
			[10, 10, 60, 40],
		],
	])(
		"outlines a %s where the browser strokes it, with square joins",
		async (_, item, box) => {
			const xml =
				(await slideXml([
					{ ...item, border: { color: "FF6347", transparency: 0, width: 4 } },
				])) ?? "";
			const [, x, y, width, height] =
				/<a:off x="(\d+)" y="(\d+)"\/><a:ext cx="(\d+)" cy="(\d+)"\/>/
					.exec(xml.slice(xml.indexOf("<p:sp>")))
					?.map(Number) ?? [];

			expect({
				box: [x, y, width, height].map((emu) => emu / 12700),
				join: /<a:ln w="50800">.*<a:miter lim="800000"\/><\/a:ln>/.test(xml),
			}).toEqual({ box, join: true });
		},
	);

	it("dashes a line along its length when its border is dashed", async () => {
		expect(
			await slideXml([
				{
					border: {
						color: "C8D7FA",
						style: "dashed",
						transparency: 0,
						width: 6,
					},
					box: { height: 0, width: 244.1, x: 10, y: 10 },
					decorative: false,
					key: "divider",
					kind: "line",
					rotation: 90,
				},
			]),
		).toMatch(
			/<a:xfrm rot="5400000">.*<a:prstGeom prst="line">.*<a:ln w="76200"><a:solidFill><a:srgbClr val="C8D7FA"\/><\/a:solidFill><a:custDash><a:ds d="200000" sp="97564"\/><\/a:custDash><a:miter lim="800000"\/><\/a:ln>/s,
		);
	});

	it.each([
		["dotted", '<a:custDash><a:ds d="100000" sp="100000"/></a:custDash>'],
		["double", '<a:ln w="50800" cmpd="dbl">'],
	] as const)(
		"outlines a box in its %s border's pattern",
		async (style, outline) => {
			expect(
				await slideXml([
					boxItem({
						border: { color: "FF6347", style, transparency: 0, width: 4 },
					}),
				]),
			).toContain(outline);
		},
	);

	it("draws no line when a box has no border", async () => {
		expect(
			await slideXml([boxItem({ fill: { color: "FF0000", transparency: 0 } })]),
		).toContain(
			'<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:ln><a:noFill/></a:ln>',
		);
	});

	it("leaves a link off its text when the text isn't underlined, as PowerPoint underlines linked text", async () => {
		expect(
			await slideXml([textItem([[run("Link", { href: "https://sli.dev" })]])]),
		).not.toContain("<a:hlinkClick");
	});

	it("keeps a link on its text when the text is underlined", async () => {
		expect(
			await slideXml([
				textItem([
					[
						run("Link", {
							href: "https://sli.dev",
							underline: { style: "single" },
						}),
					],
				]),
			]),
		).toMatch(/<a:rPr [^>]*u="sng"[^>]*>.*<a:hlinkClick r:id="rId\d+"/s);
	});

	it("keeps a link on its text when it's code, which gets no hotspot", async () => {
		expect(
			await slideXml([
				{
					baselines: [40],
					box: { height: 40, width: 400, x: 40, y: 20 },
					decorative: false,
					key: "code",
					kind: "code",
					paragraphs: [
						{ runs: [run("link", { font: "Mono", href: "https://sli.dev" })] },
					],
					style: { ...style, nowrap: true },
				},
			]),
		).toContain("<a:hlinkClick");
	});

	it("links a hotspot box to its target", async () => {
		const zip = await writeItems([
			boxItem({
				fill: { color: "FFFFFF", transparency: 100 },
				href: "https://sli.dev",
			}),
		]);

		expect({
			rels: (await part(zip, "ppt/slides/_rels/slide1.xml.rels"))?.match(
				/Target="https:\/\/sli\.dev"/g,
			),
			shape: (await part(zip, "ppt/slides/slide1.xml"))?.match(
				/<p:cNvPr [^>]*><a:hlinkClick r:id="rId\d+"/g,
			)?.length,
		}).toEqual({ rels: ['Target="https://sli.dev"'], shape: 1 });
	});

	it("underlines a run when it's underlined", async () => {
		expect(
			await slideXml([
				textItem([[run("Under", { underline: { style: "single" } })]]),
			]),
		).toContain('<a:rPr lang="en-US" kern="100" sz="2400" u="sng" dirty="0">');
	});

	it("underlines a run in its underline's style and color when it has them", async () => {
		expect(
			await slideXml([
				textItem([
					[run("Wavy", { underline: { color: "FF0000", style: "wavy" } })],
				]),
			]),
		).toContain(
			'<a:rPr lang="en-US" kern="100" sz="2400" u="wavy" dirty="0"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:uFill><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:uFill>',
		);
	});

	it("writes paragraphs right to left when their text runs right to left", async () => {
		expect(
			await slideXml([
				textItem([[run("Right to left")]], {
					style: { ...style, align: "right", nowrap: false, rtl: true },
				}),
			]),
		).toContain('<a:pPr rtl="1"  algn="r"');
	});

	it("rotates a box when it's rotated", async () => {
		expect(
			await slideXml([
				boxItem({ box: { height: 10, width: 20, x: 0, y: 0 }, rotation: 45 }),
			]),
		).toContain('<a:xfrm rot="2700000">');
	});

	it("strikes a run through when it's struck", async () => {
		expect(
			await slideXml([textItem([[run("Gone", { strike: true })]])]),
		).toContain('<a:rPr lang="en-US" kern="100" sz="2400" strike="sngStrike"');
	});

	it("capitalizes a run without its stand-in language when it's uppercase", async () => {
		expect(
			await slideXml([textItem([[run("Loud", { upper: true })]])]),
		).toContain(
			'<a:rPr lang="en-US" kern="100" cap="all" sz="2400" dirty="0">',
		);
	});

	it("kerns a run when it's letter-spaced", async () => {
		expect(
			await slideXml([textItem([[run("Wide", { letterSpacing: 2 })]])]),
		).toContain(
			'<a:rPr lang="en-US" sz="2400" spc="200" kern="100" dirty="0">',
		);
	});

	it("keeps a real color when it's one a gradient placeholder would use", async () => {
		const xml = await slideXml([
			boxItem({ fill: { color: "FEA000", transparency: 0 }, key: "solid" }),
			boxItem({
				box: { height: 10, width: 10, x: 20, y: 0 },
				fill: {
					gradient: {
						angle: 90,
						stops: [
							{ color: "FF0000", position: 0, transparency: 0 },
							{ color: "0000FF", position: 100, transparency: 50 },
						],
					},
				},
				key: "gradient",
			}),
		]);

		expect(
			xml?.match(/<a:(?:solidFill|gradFill)\b.*?<\/a:(?:solidFill|gradFill)>/g),
		).toEqual([
			'<a:solidFill><a:srgbClr val="FEA000"/></a:solidFill>',
			'<a:gradFill flip="none" rotWithShape="1"><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"></a:srgbClr></a:gs><a:gs pos="100000"><a:srgbClr val="0000FF"><a:alpha val="50000"/></a:srgbClr></a:gs></a:gsLst><a:lin ang="0" scaled="0"/></a:gradFill>',
		]);
	});

	it.each([
		["decimal", 'indent="-350520"'],
		["lower-roman", 'indent="-853440"'],
	])("leaves hanging room for the widest %s number", async (marker, indent) => {
		expect(await slideXml([numberedList([7, 8], marker)])).toContain(indent);
	});

	it("draws keyword markers in Arial when a list has them", async () => {
		expect(await slideXml([numberedList([1], "disc")])).toContain(
			'<a:buSzPct val="60000"/><a:buFont typeface="Arial"/><a:buChar char="&#x25CF;"/>',
		);
	});

	it("writes no marker or gutter when a list's marker is an empty string", async () => {
		expect(await slideXml([numberedList([1], '""')])).toContain(
			'<a:pPr algn="l" marL="127" indent="0"><a:lnSpc><a:spcPct val="125000"/></a:lnSpc><a:buNone/></a:pPr>',
		);
	});

	it("styles each paragraph's marker when a disc list nests a numbered one", async () => {
		const xml = await slideXml([
			listItem([
				listParagraph("Outer"),
				listParagraph("Inner", { indent: 24, level: 1, marker: "decimal" }),
			]),
		]);

		expect(xml?.match(/<a:pPr\b.*?<\/a:pPr>/g)).toEqual([
			'<a:pPr algn="l" marL="228600" indent="-228600"><a:lnSpc><a:spcPct val="125000"/></a:lnSpc><a:buSzPct val="60000"/><a:buFont typeface="Arial"/><a:buChar char="&#x25CF;"/></a:pPr>',
			// The number hangs by its own width, one digit and a period at 24px, not the outer disc's.
			'<a:pPr algn="l" lvl="1" marL="533400" indent="-350520"><a:lnSpc><a:spcPct val="125000"/></a:lnSpc><a:buSzPct val="100000"/><a:buFontTx/><a:buAutoNum type="arabicPeriod" startAt="1"/></a:pPr>',
		]);
	});

	it("hangs each bullet by its own marker's size", async () => {
		const xml = await slideXml([
			listItem([
				listParagraph("Small", { markerFace: { font: "Body", size: 12 } }),
				listParagraph("Large", { markerFace: { font: "Body", size: 24 } }),
			]),
		]);

		// 0.75em of 12px and 24px.
		expect(
			[...(xml ?? "").matchAll(/ indent="(-\d+)"/g)].map(
				([, indent]) => indent,
			),
		).toEqual(["-114300", "-228600"]);
	});

	it("draws a nested keyword marker in Arial when a numbered list has one", async () => {
		const xml = await slideXml([
			listItem([
				listParagraph("Outer", { marker: "decimal" }),
				listParagraph("Inner", { indent: 24, level: 1, marker: "circle" }),
			]),
		]);

		expect(xml?.match(/<a:bu\w.*?(?=<\/a:pPr>)/g)).toEqual([
			'<a:buSzPct val="100000"/><a:buFontTx/><a:buAutoNum type="arabicPeriod" startAt="1"/>',
			'<a:buSzPct val="60000"/><a:buFont typeface="Arial"/><a:buChar char="&#x25CB;"/>',
		]);
	});

	it("colors only shown markers when a list's markers have a color", async () => {
		const xml = await slideXml([
			listItem([
				listParagraph("Shown", { markerColor: "FF0000" }),
				listParagraph("Continued", { bullet: false, markerColor: "FF0000" }),
			]),
		]);

		expect(xml?.match(/<a:bu\w.*?(?=<\/a:pPr>)/g)).toEqual([
			'<a:buClr><a:srgbClr val="FF0000"/></a:buClr><a:buSzPct val="60000"/><a:buFont typeface="Arial"/><a:buChar char="&#x25CF;"/>',
			"<a:buNone/>",
		]);
	});

	it("colors a string marker in the text's font when a list has one", async () => {
		expect(
			await slideXml([
				listItem([
					listParagraph("Arrow", { marker: '"→"', markerColor: "FF0000" }),
				]),
			]),
		).toContain(
			'<a:buClr><a:srgbClr val="FF0000"/></a:buClr><a:buSzPct val="100000"/><a:buChar char="&#x2192;"/>',
		);
	});

	it.each([
		// PowerPoint takes a marker's color from its paragraph's first run.
		["matches its first run", [run("Text")], "DAE7F3", undefined],
		["differs from its first run", [run("Text")], "FF0000", "FF0000"],
		[
			"follows a spacer",
			[run("  ", { pad: true }), run("Text")],
			"DAE7F3",
			"DAE7F3",
		],
		[
			"follows a hidden run",
			[run("Hidden", { transparency: 100 }), run("Text")],
			"DAE7F3",
			"DAE7F3",
		],
	] as const)(
		"colors each marker only where it needs to when its color %s",
		async (_, runs, markerColor, written) => {
			const xml = await slideXml([
				listItem([
					{ ...listParagraph("First", { markerColor }), runs: [...runs] },
					listParagraph("Second", { markerColor: "00FF00" }),
				]),
			]);

			expect(
				xml
					?.match(/<a:pPr\b.*?<\/a:pPr>/g)
					?.map((pPr) => /<a:buClr><a:srgbClr val="(\w+)"/.exec(pPr)?.[1]),
			).toEqual([written, "00FF00"]);
		},
	);

	it("draws a numbered marker in its item's font and size when its first run's differ", async () => {
		const xml = await slideXml([
			listItem([
				listParagraph("", {
					marker: "decimal",
					markerFace: { font: "Body", size: 24 },
					runs: [run("code", { font: "Mono", size: 20 }), run(" first")],
				}),
			]),
		]);

		expect(xml?.match(/<a:bu\w.*?(?=<\/a:pPr>)/g)).toEqual([
			'<a:buSzPts val="2400"/><a:buFont typeface="Body"/><a:buAutoNum type="arabicPeriod" startAt="1"/>',
		]);
	});

	it("sizes a keyword marker from its item's size when its first run's differs", async () => {
		const xml = await slideXml([
			listItem([
				listParagraph("", {
					markerFace: { font: "Body", size: 24 },
					runs: [run("code", { font: "Mono", size: 20 }), run(" first")],
				}),
			]),
		]);

		expect(xml?.match(/<a:bu\w.*?(?=<\/a:pPr>)/g)).toEqual([
			'<a:buSzPts val="1440"/><a:buFont typeface="Arial"/><a:buChar char="&#x25CF;"/>',
		]);
	});

	it("draws a string marker in its item's font when its first run's differs", async () => {
		const xml = await slideXml([
			listItem([
				listParagraph("", {
					marker: '"→"',
					markerFace: { font: "Body", size: 24 },
					runs: [run("code", { font: "Mono", size: 24 }), run(" first")],
				}),
			]),
		]);

		expect(xml?.match(/<a:bu\w.*?(?=<\/a:pPr>)/g)).toEqual([
			'<a:buSzPts val="2400"/><a:buFont typeface="Body"/><a:buChar char="&#x2192;"/>',
		]);
	});

	it("leaves a marker in its first run's font and size when they match its item's", async () => {
		const xml = await slideXml([
			listItem([
				listParagraph("Same", {
					marker: "decimal",
					markerFace: { font: "Body", size: 24 },
				}),
			]),
		]);

		expect(xml?.match(/<a:bu\w.*?(?=<\/a:pPr>)/g)).toEqual([
			'<a:buSzPct val="100000"/><a:buFontTx/><a:buAutoNum type="arabicPeriod" startAt="1"/>',
		]);
	});

	it("hangs a string marker by its text's width when its font is loaded", async () => {
		expect(
			await slideXml([
				listItem([
					listParagraph("Arrow", {
						marker: '"→ "',
						runs: [run("Arrow", { font: "Measured" })],
					}),
				]),
			]),
		).toContain('<a:pPr algn="l" marL="304800" indent="-304800">');
	});

	it("hangs each decimal marker by its own width, leaving room for the widest, when its font is loaded", async () => {
		const xml = await slideXml([
			listItem(
				[9, 10].map((number) =>
					listParagraph(`Item ${String(number)}`, {
						marker: "decimal",
						number,
						runs: [run(`Item ${String(number)}`, { font: "Measured" })],
					}),
				),
			),
		]);

		// "9. " and "10. " at 0.5em a character, at 24px.
		expect(
			[...(xml ?? "").matchAll(/marL="(\d+)" indent="(-\d+)"/g)].map(
				([, marL, indent]) => [marL, indent],
			),
		).toEqual([
			["609600", "-457200"],
			["609600", "-609600"],
		]);
	});

	it("spaces each list item by its own line height", async () => {
		const xml = await slideXml([
			listItem([
				listParagraph("Loose"),
				listParagraph("Single", { lineHeight: 28.8 }),
			]),
		]);

		expect(
			[...(xml ?? "").matchAll(/<a:spcPct val="(\d+)"\/>/g)].map(
				([, value]) => value,
			),
		).toEqual(["125000", "100000"]);
	});

	it.each([
		// From 9.14px below a 36px line's baseline, down 18.24px of a single-spaced 24px line, leaves 2.6pt of the 30px step.
		[36, 24, 30, 300],
		// From 5.76px below a 24px line's baseline, down 26.86px of a 36px line, leaves 2.4pt of the 35px step.
		[24, 36, 35, 200],
	])(
		"spaces a list item after one with a %spx pitch, before one with %spx, by what's left of a %spx step",
		async (from, to, step, spaceAfter) => {
			const xml = await slideXml([
				listItem([
					listParagraph("First", { baselines: [40], lineHeight: from }),
					listParagraph("Second", {
						baselines: [40 + step],
						lineHeight: to,
					}),
				]),
			]);

			expect(xml).toContain(
				`<a:spcAft><a:spcPts val="${String(spaceAfter)}"/></a:spcAft>`,
			);
		},
	);

	it("makes up for a step PowerPoint can't space as short with less space later", async () => {
		const xml = await slideXml([
			listItem(
				[36, 24, 36, 24].map((lineHeight, index) =>
					listParagraph(`Item ${String(index)}`, {
						baselines: [40 + 30 * index],
						lineHeight,
					}),
				),
			),
		]);

		// Each pair of steps is 60px, as PowerPoint's are: 2.6pt after the first leaves it 2.6pt ahead.
		expect(
			[...(xml ?? "").matchAll(/<a:spcAft><a:spcPts val="(\d+)"/g)].map(
				([, value]) => value,
			),
		).toEqual(["300"]);
	});

	it("nests a paragraph a level deeper when its list is nested", async () => {
		expect(
			await slideXml([
				listItem([
					listParagraph("Parent"),
					listParagraph("Child", { indent: 24, level: 1 }),
				]),
			]),
		).toContain('<a:pPr algn="l" lvl="1" marL="533400" indent="-228600">');
	});

	it("hangs a marker just before its text when the item itself is indented", async () => {
		expect(
			await slideXml([listItem([listParagraph("Indented", { indent: 60 })])]),
		).toContain('<a:pPr algn="l" marL="990600" indent="-228600">');
	});

	it("numbers an item from its own number when a nested list comes before it", async () => {
		expect(
			await slideXml([
				listItem([
					listParagraph("First", { marker: "decimal" }),
					listParagraph("Inner", { indent: 24, level: 1 }),
					listParagraph("Second", { marker: "decimal", number: 2 }),
				]),
			]),
		).toContain('<a:buAutoNum type="arabicPeriod" startAt="2"/>');
	});

	it("numbers a list from 1 and reports it when the list starts at 0", async () => {
		const { file, result } = await write(
			deckOf([
				{ states: [{ clicks: 0, items: [numberedList([0, 1], "decimal")] }] },
			]),
		);
		const xml = await part(
			await JSZip.loadAsync(file),
			"ppt/slides/slide1.xml",
		);

		expect({
			problems: result.problems,
			starts: xml?.match(/startAt="\d+"/g),
		}).toEqual({
			problems: [
				{
					message:
						"PowerPoint can't number list items below 1, so a list here is numbered from 1.",
					severity: "warning",
					slide: 1,
				},
			],
			starts: ['startAt="1"', 'startAt="1"'],
		});
	});

	it("indents a paragraph without a marker when it continues an item", async () => {
		expect(
			await slideXml([
				listItem([
					listParagraph("First"),
					listParagraph("More", { bullet: false }),
				]),
			]),
		).toContain(
			'<a:pPr algn="l" marL="228600" indent="0"><a:lnSpc><a:spcPct val="125000"/></a:lnSpc><a:buNone/></a:pPr><a:r><a:rPr lang="en-US" kern="100" sz="2400" dirty="0"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:rPr><a:t>More</a:t>',
		);
	});

	it("hides a list's paragraph when it's hidden in every click state", async () => {
		expect(
			await slideXml([
				listItem([
					listParagraph("Shown"),
					listParagraph("Hidden", { visible: false }),
				]),
			]),
		).toContain(
			'<a:buNone/></a:pPr><a:r><a:rPr lang="en-US" kern="100" sz="2400" dirty="0"><a:solidFill><a:schemeClr val="tx1"><a:alpha val="0"/></a:schemeClr></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:rPr><a:t>Hidden</a:t>',
		);
	});

	it("lowers a run when it's a subscript", async () => {
		expect(
			await slideXml([textItem([[run("H"), run("2", { script: "sub" })]])]),
		).toContain('baseline="-40000"');
	});

	it("keeps a shadow's zero offset when text glows", async () => {
		expect(
			await slideXml([
				textItem([
					[
						run("Glow", {
							shadow: {
								blur: 8,
								color: "FFFFFF",
								offsetX: 0,
								offsetY: 0,
								transparency: 0,
							},
						}),
					],
				]),
			]),
		).toContain('blurRad="101600" dist="127"');
	});

	it("casts a shadow to the right when it's offset straight right", async () => {
		expect(
			await slideXml([
				textItem([
					[
						run("Right", {
							shadow: {
								blur: 0,
								color: "000000",
								offsetX: 4,
								offsetY: 0,
								transparency: 0,
							},
						}),
					],
				]),
			]),
		).toContain('dist="50800" dir="21540000"');
	});

	it("casts no shadow when it's transparent", async () => {
		expect(
			await slideXml([
				textItem([
					[
						run("Faint", {
							shadow: {
								blur: 4,
								color: "000000",
								offsetX: 2,
								offsetY: 2,
								transparency: 100,
							},
						}),
					],
				]),
			]),
		).not.toContain("<a:outer");
	});

	it("casts a box's shadow when it has one", async () => {
		expect(
			await slideXml([
				boxItem({
					shadow: {
						blur: 8,
						color: "000000",
						offsetX: 0,
						offsetY: 4,
						transparency: 75,
					},
				}),
			]),
		).toContain(
			'rotWithShape="0" blurRad="101600" dist="50800" dir="5400000"> <a:srgbClr val="000000"> <a:alpha val="25000"/></a:srgbClr>',
		);
	});

	it("casts a picture's shadow when it has one", async () => {
		expect(
			await slideXml([
				imageItem({
					shadow: {
						blur: 8,
						color: "000000",
						offsetX: 0,
						offsetY: 4,
						transparency: 75,
					},
				}),
			]),
		).toContain(
			'rotWithShape="0" blurRad="101600" dist="50800" dir="5400000"><a:srgbClr val="000000"><a:alpha val="25000"/></a:srgbClr>',
		);
	});

	it("joins whitespace into its neighboring runs when only its color differs", async () => {
		const red = { color: "FF0000" };

		expect(
			(
				await slideXml([
					textItem([
						[
							run("  "),
							run("const", red),
							run(" "),
							run("a", red),
							run(" ", { underline: { style: "single" } }),
						],
					]),
				])
			)?.match(/<a:t>[^<]*<\/a:t>/g),
		).toEqual(["<a:t>  const a</a:t>", "<a:t> </a:t>"]);
	});

	it("places rotated text by its line box, turned about the same center, when it's rotated", async () => {
		const [, y, height] =
			/<a:xfrm rot="5400000"><a:off x="508000" y="(\d+)"\/><a:ext cx="5080000" cy="(\d+)"\/>/
				.exec(
					(await slideXml([
						textItem([[run("Turned")]], { baselines: undefined, rotation: 90 }),
					])) ?? "",
				)
				?.map(Number) ?? [];

		// The line's baseline is half its 12px leading, plus the 19px ascent Blink rounds 19.2 to, below its top, and PowerPoint's 0.746 down its 36px pitch.
		expect(y / 12700).toBeCloseTo(20 + 6 + 19 - 0.746 * 36, 2);
		// The 40px box's center.
		expect((y + height / 2) / 12700).toBeCloseTo(40, 2);
	});

	it("makes a slide's first heading its title", async () => {
		expect(
			await slideXml([
				textItem([[run("Body")]], { key: "body" }),
				textItem([[run("Heading")]], { heading: true, key: "a" }),
				textItem([[run("Later")]], { heading: true, key: "b" }),
			]),
		).toMatch(
			/name="Text: Heading"><\/p:cNvPr><p:cNvSpPr><a:spLocks noGrp="1"\/><\/p:cNvSpPr><p:nvPr><p:ph type="title"\/><\/p:nvPr>.*name="Text: Later"><\/p:cNvPr><p:cNvSpPr txBox="1"\/><p:nvPr><\/p:nvPr>/s,
		);
	});

	it("marks a shape decorative when it's decoration", async () => {
		expect(await slideXml([boxItem({ decorative: true })])).toContain(
			'name="Box"><a:extLst><a:ext uri="{C183D7F6-B498-43B3-948B-1728B52AA6E4}"><adec:decorative xmlns:adec="http://schemas.microsoft.com/office/drawing/2017/decorative" val="1"/></a:ext></a:extLst></p:cNvPr>',
		);
	});

	it("marks a picture decorative without a description when it has no alt text", async () => {
		expect(await slideXml([imageItem({ alt: "", decorative: true })])).toMatch(
			/name="Image" descr="">\s*<a:extLst><a:ext uri="\{C183D7F6-B498-43B3-948B-1728B52AA6E4\}">/,
		);
	});

	it("outlines stroked text over its fill when CSS paints the stroke last", async () => {
		const xml = await slideXml([
			textItem([
				[
					run("Inked", {
						stroke: { behind: false, color: "000000", width: 2 },
					}),
				],
			]),
		]);

		expect(xml).toMatch(/name="Text: Inked">.*?<a:rPr[^>]*><a:ln w=/s);
		expect(xml).not.toContain("(outline)");
	});

	it("marks a text stroke's outlined copy decorative when CSS paints the stroke behind the fill", async () => {
		expect(
			await slideXml([
				textItem([
					[
						run("Inked", {
							stroke: { behind: true, color: "000000", width: 2 },
						}),
					],
				]),
			]),
		).toMatch(
			/name="Text: Inked \(outline\)"><a:extLst>.*?<adec:decorative .*?name="Text: Inked"><\/p:cNvPr>(?:(?!<a:ln w=).)*<\/p:sp>/s,
		);
	});

	it("flips and turns a picture when it's flipped and rotated", async () => {
		expect(await slideXml([imageItem({ flip: true, rotation: -90 })])).toMatch(
			/<p:pic>.*<a:xfrm(?=[^>]*flipH="1")(?=[^>]*rot="16200000")/s,
		);
	});

	it("flips a box's fill when it's flipped", async () => {
		expect(await slideXml([boxItem({ flip: true })])).toMatch(
			/<p:sp>.*<a:xfrm(?=[^>]*flipH="1")/s,
		);
	});

	it("rounds a picture's corners by its shorter side", async () => {
		expect(await slideXml([imageItem({ radius: 20 })])).toContain(
			'<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 20000"/></a:avLst></a:prstGeom>',
		);
	});

	it("rounds a picture into an ellipse when its radius reaches half its shorter side", async () => {
		expect(await slideXml([imageItem({ radius: 50 })])).toMatch(
			/<p:pic>(?:(?!<\/p:pic>).)*<a:prstGeom prst="ellipse">/s,
		);
	});

	it("points identical pictures at one media file when their files differ", async () => {
		const copies = ["a.png", "b.png"].map((name) => path.join(directory, name));
		for (const copy of copies) {
			await fs.writeFile(
				copy,
				Buffer.concat([await fs.readFile(photo), Buffer.from([2])]),
			);
		}

		const zip = await writeItems([
			imageItem({ file: copies[0], key: "a" }),
			imageItem({ file: copies[1], key: "b" }),
		]);

		expect({
			media: zip.file(/^ppt\/media\/image/).map(({ name }) => name),
			rels: (await part(zip, "ppt/slides/_rels/slide1.xml.rels"))?.match(
				/Target="[^"]*media[^"]*"/g,
			),
		}).toEqual({
			media: ["ppt/media/image-1-1.png"],
			rels: [
				'Target="../media/image-1-1.png"',
				'Target="../media/image-1-1.png"',
			],
		});
	});

	it("declares JPEGs by their standard type when the deck has them", async () => {
		const jpeg = path.join(directory, "photo.jpg");
		await fs.writeFile(
			jpeg,
			Buffer.concat([await fs.readFile(photo), Buffer.from([1])]),
		);

		expect(
			await part(
				await writeItems([imageItem({ file: jpeg })]),
				"[Content_Types].xml",
			),
		).toMatch(/^(?!.*image\/jpg").*Extension="jpg" ContentType="image\/jpeg"/s);
	});

	it("points a link's runs at one relationship when a line break splits it", async () => {
		const href = "https://sli.dev";
		const underline = { style: "single" } as const;
		const zip = await writeItems([
			textItem([
				[
					run("first", { href, underline }),
					run("second", { href, lineBreak: "forced", underline }),
				],
			]),
		]);

		expect({
			rels: (await part(zip, "ppt/slides/_rels/slide1.xml.rels"))?.match(
				/Target="https:\/\/sli\.dev"/g,
			),
			runs: (await part(zip, "ppt/slides/slide1.xml"))?.match(/r:id="rId\d+"/g),
		}).toEqual({
			rels: ['Target="https://sli.dev"'],
			runs: ['r:id="rId1"', 'r:id="rId1"'],
		});
	});

	const themedItems = [
		textItem([[run("Heading", { font: "Heading" })]], {
			heading: true,
			key: "heading",
		}),
		textItem(
			[
				[
					run("Most of the body text"),
					run("a link", { color: "3366CC", href: "https://sli.dev" }),
				],
			],
			{ key: "body" },
		),
	];

	it("themes text, background, and links by the colors the deck uses most", async () => {
		expect(
			(
				await part(await writeItems(themedItems), "ppt/theme/theme1.xml")
			)?.match(/<a:clrScheme\b.*?<\/a:clrScheme>/)?.[0],
		).toBe(
			'<a:clrScheme name="Deck Title"><a:dk1><a:srgbClr val="101010"/></a:dk1><a:lt1><a:srgbClr val="DAE7F3"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="3366CC"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="3366CC"/></a:hlink><a:folHlink><a:srgbClr val="3366CC"/></a:folHlink></a:clrScheme>',
		);
	});

	it("themes fonts by what headings and body text use most", async () => {
		expect(
			(
				await part(await writeItems(themedItems), "ppt/theme/theme1.xml")
			)?.match(
				/<a:(?:major|minor)Font><a:latin [^>]*\/><a:ea [^>]*\/><a:cs [^>]*\/>/g,
			),
		).toEqual([
			'<a:majorFont><a:latin typeface="Heading"/><a:ea typeface="Heading"/><a:cs typeface="Heading"/>',
			'<a:minorFont><a:latin typeface="Body"/><a:ea typeface="Body"/><a:cs typeface="Body"/>',
		]);
	});

	it("picks accents by how many shapes and texts use them when the deck colors several", async () => {
		const theme = await part(
			await writeItems([
				boxItem({ fill: { color: "00A3FF", transparency: 0 }, key: "a" }),
				boxItem({ fill: { color: "00A3FF", transparency: 0 }, key: "b" }),
				boxItem({
					border: { color: "55E1F3", transparency: 0, width: 1 },
					key: "c",
				}),
				boxItem({ fill: { color: "FFFFFF", transparency: 97 }, key: "d" }),
			]),
			"ppt/theme/theme1.xml",
		);

		expect(theme).toContain(
			'<a:accent1><a:srgbClr val="00A3FF"/></a:accent1><a:accent2><a:srgbClr val="55E1F3"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>',
		);
	});

	it("refers runs to the theme's fonts and colors when they're the theme's", async () => {
		expect(
			await slideXml([
				textItem([
					[
						run("Themed"),
						run("Linked", {
							href: "https://sli.dev",
							underline: { style: "single" },
						}),
					],
				]),
			]),
		).toContain(
			'<a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:rPr><a:t>Themed</a:t></a:r><a:r><a:rPr lang="en-US" kern="100" sz="2400" u="sng" dirty="0"><a:solidFill><a:schemeClr val="hlink"/></a:solidFill>',
		);
	});

	it("styles the master's title like the slides' titles when it writes it", async () => {
		const master = await part(
			await writeItems([
				textItem([[run("Title", { color: "E5EDF3", font: "Heading" })]], {
					heading: true,
				}),
				textItem([[run("Body text")]], { key: "body" }),
			]),
			"ppt/slideMasters/slideMaster1.xml",
		);

		expect({
			placeholder: master?.match(
				/name="Title Placeholder 1".*?<\/p:spPr>/,
			)?.[0],
			style: firstLevel(master, "p:titleStyle"),
		}).toEqual({
			placeholder:
				'name="Title Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="508000" y="166929"/><a:ext cx="5130800" cy="595071"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>',
			style: {
				align: "l",
				properties:
					'<a:lnSpc><a:spcPct val="125000"/></a:lnSpc><a:spcBef><a:spcPct val="0"/></a:spcBef><a:buNone/><a:defRPr sz="2400" kern="1200"><a:solidFill><a:schemeClr val="tx2"/></a:solidFill><a:latin typeface="+mj-lt"/><a:ea typeface="+mj-ea"/><a:cs typeface="+mj-cs"/></a:defRPr>',
			},
		});
	});

	it("gives the master's title style the titles' letter spacing, weight, slant, and capitals", async () => {
		expect(
			firstLevel(
				await part(
					await writeItems([
						textItem(
							[
								[
									run("Title", {
										bold: true,
										italic: true,
										letterSpacing: -1,
										upper: true,
									}),
								],
							],
							{ heading: true },
						),
					]),
					"ppt/slideMasters/slideMaster1.xml",
				),
				"p:titleStyle",
			)?.properties,
		).toContain(
			'<a:defRPr sz="2400" b="1" i="1" cap="all" spc="-100" kern="1200">',
		);
	});

	it("colors titles with the theme's second text color when their color isn't the text's", async () => {
		expect(
			(
				await part(
					await writeItems([
						textItem([[run("Title", { color: "E5EDF3" })]], { heading: true }),
						textItem([[run("Body text")]], { key: "body" }),
					]),
					"ppt/theme/theme1.xml",
				)
			)?.match(/<a:lt2>.*?<\/a:lt2>/)?.[0],
		).toBe('<a:lt2><a:srgbClr val="E5EDF3"/></a:lt2>');
	});

	it("leaves a title's runs to the master's title style when they look like it", async () => {
		expect(
			await slideXml([textItem([[run("Title")]], { heading: true })]),
		).toContain(
			'<p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="508000" y="166929"/><a:ext cx="5130800" cy="595071"/></a:xfrm><a:prstGeom prst="rect"><a:avLst></a:avLst></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0" anchor="t"></a:bodyPr><a:lstStyle/><a:p><a:pPr algn="l" indent="0" marL="0"><a:lnSpc><a:spcPct val="125000"/></a:lnSpc><a:buNone/></a:pPr><a:r><a:rPr lang="en-US" kern="100" dirty="0"/><a:t>Title</a:t></a:r><a:endParaRPr lang="en-US" kern="100" dirty="0"/></a:p>',
		);
	});

	it("keeps a title run's own look and spells out what the master would change when the run looks different", async () => {
		expect(
			await slideXml([
				textItem(
					[
						[
							run("Tight", { letterSpacing: -1 }),
							run("code", { font: "Mono" }),
						],
					],
					{
						heading: true,
					},
				),
			]),
		).toContain(
			'<a:r><a:rPr lang="en-US" kern="100" dirty="0" spc="0"><a:latin typeface="Mono" pitchFamily="34" charset="0"/><a:ea typeface="Mono" pitchFamily="34" charset="-122"/><a:cs typeface="Mono" pitchFamily="34" charset="-120"/></a:rPr><a:t>code</a:t></a:r>',
		);
	});

	it("names a title font in the master when it isn't one of the theme's", async () => {
		expect(
			firstLevel(
				await part(
					await writeItems([
						textItem(
							[
								[
									run("A", { font: "Display" }),
									run(" much longer heading", { font: "Heading" }),
								],
							],
							{ heading: true },
						),
						textItem([[run("Body text")]], { key: "body" }),
					]),
					"ppt/slideMasters/slideMaster1.xml",
				),
				"p:titleStyle",
			)?.properties,
		).toContain(
			'<a:latin typeface="Display"/><a:ea typeface="Display"/><a:cs typeface="Display"/>',
		);
	});

	it("gives new shapes the look of the deck's most common box", async () => {
		const card = {
			border: { color: "336699", transparency: 90, width: 1 },
			fill: { color: "336699", transparency: 97.5 },
		};

		expect(
			(
				await part(
					await writeItems([
						boxItem({ ...card, key: "a" }),
						boxItem({ ...card, key: "b" }),
						boxItem({ fill: { color: "FF0000", transparency: 0 }, key: "c" }),
					]),
					"ppt/theme/theme1.xml",
				)
			)?.match(/<a:objectDefaults>.*<\/a:objectDefaults>/)?.[0],
		).toBe(
			'<a:objectDefaults><a:spDef><a:spPr><a:solidFill><a:srgbClr val="336699"><a:alpha val="2500"/></a:srgbClr></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="336699"><a:alpha val="10000"/></a:srgbClr></a:solidFill></a:ln></a:spPr><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/><a:style><a:lnRef idx="0"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef></a:style></a:spDef></a:objectDefaults>',
		);
	});

	it("gives new text the deck's most used body size and spacing", async () => {
		const zip = await writeItems([
			textItem([[run("Body text")]]),
			textItem([[run("Rare", { size: 12 })]], { key: "rare" }),
		]);
		const body = {
			align: "l",
			properties:
				'<a:lnSpc><a:spcPct val="125000"/></a:lnSpc><a:defRPr sz="2400" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr>',
		};

		expect({
			master: firstLevel(
				await part(zip, "ppt/slideMasters/slideMaster1.xml"),
				"p:otherStyle",
			),
			presentation: firstLevel(
				await part(zip, "ppt/presentation.xml"),
				"p:defaultTextStyle",
			),
		}).toEqual({ master: body, presentation: body });
	});

	it("maps text to the light color when the background is dark", async () => {
		const zip = await writeItems([textItem([[run("Light")]])]);

		expect(await part(zip, "ppt/slideMasters/slideMaster1.xml")).toContain(
			'<p:clrMap bg1="dk1" tx1="lt1" bg2="dk2" tx2="lt2"',
		);
	});

	it("writes one Title Only layout", async () => {
		const zip = await writeItems([]);

		expect({
			layout: await part(zip, "ppt/slideLayouts/slideLayout1.xml"),
			layouts: zip.file(/^ppt\/slideLayouts\/[^/]+$/).map(({ name }) => name),
		}).toEqual({
			layout: expect.stringMatching(
				/type="titleOnly" preserve="1">\s*<p:cSld name="Title Only"><p:spTree>/,
			),
			layouts: ["ppt/slideLayouts/slideLayout1.xml"],
		});
	});

	it("puts the background on the master", async () => {
		expect(
			await part(await writeItems([]), "ppt/slideMasters/slideMaster1.xml"),
		).toContain(
			'<p:cSld><p:bg><p:bgPr><a:blipFill dpi="0" rotWithShape="1"><a:blip r:embed="rId3"/>',
		);
	});

	it("keeps the notes master on pptxgenjs's own theme when it themes the slides", async () => {
		const zip = await writeItems([textItem([[run("Text")]])]);

		expect({
			contentTypes: await part(zip, "[Content_Types].xml"),
			notesMaster: await part(
				zip,
				"ppt/notesMasters/_rels/notesMaster1.xml.rels",
			),
			notesTheme: await part(zip, "ppt/theme/theme2.xml"),
		}).toEqual({
			contentTypes: expect.stringMatching(
				/<Override PartName="\/ppt\/theme\/theme2\.xml" ContentType="[^"]+theme\+xml"\/>/,
			),
			notesMaster: expect.stringContaining('Target="../theme/theme2.xml"'),
			notesTheme: expect.stringContaining(
				'<a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>',
			),
		});
	});

	it("lists the notes master before the slides", async () => {
		expect(await part(await writeItems([]), "ppt/presentation.xml")).toMatch(
			/<\/p:sldMasterIdLst><p:notesMasterIdLst>.*?<\/p:notesMasterIdLst><p:sldIdLst>/,
		);
	});

	it("sizes notes pages as portrait pages", async () => {
		expect(await part(await writeItems([]), "ppt/presentation.xml")).toContain(
			'<p:notesSz cx="6858000" cy="9144000"/>',
		);
	});

	it("lists the deck's fonts, theme, and slide titles in its properties", async () => {
		const app = await part(
			await writeZip(
				deckOf([
					{
						states: [
							{
								clicks: 0,
								items: [
									textItem([[run("Hello & welcome")]], { heading: true }),
								],
							},
						],
					},
					{ states: [{ clicks: 0, items: [] }] },
				]),
			),
			"docProps/app.xml",
		);

		expect({
			headings: app?.match(/<HeadingPairs>.*<\/HeadingPairs>/s)?.[0],
			titles: app?.match(/<TitlesOfParts>.*<\/TitlesOfParts>/s)?.[0],
		}).toEqual({
			headings:
				'<HeadingPairs><vt:vector size="6" baseType="variant"><vt:variant><vt:lpstr>Fonts Used</vt:lpstr></vt:variant><vt:variant><vt:i4>2</vt:i4></vt:variant><vt:variant><vt:lpstr>Theme</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant><vt:variant><vt:lpstr>Slide Titles</vt:lpstr></vt:variant><vt:variant><vt:i4>2</vt:i4></vt:variant></vt:vector></HeadingPairs>',
			titles:
				'<TitlesOfParts><vt:vector size="5" baseType="lpstr"><vt:lpstr>Body</vt:lpstr><vt:lpstr>Mono</vt:lpstr><vt:lpstr>Deck Title</vt:lpstr><vt:lpstr>Hello &amp; welcome</vt:lpstr><vt:lpstr>PowerPoint Presentation</vt:lpstr></vt:vector></TitlesOfParts>',
		});
	});

	it("writes no folder entries or company", async () => {
		const zip = await writeItems([]);

		expect({
			app: await part(zip, "docProps/app.xml"),
			folders: Object.values(zip.files).filter(({ dir }) => dir),
		}).toEqual({
			app: expect.stringContaining("<Company></Company>"),
			folders: [],
		});
	});

	it("formats notes' Markdown when a note is one line", async () => {
		expect(await notesXml("Say **this** and `that`")).toContain(
			'<a:p><a:r><a:rPr lang="en-US" dirty="0"/><a:t>Say </a:t></a:r><a:r><a:rPr lang="en-US" b="1" dirty="0"/><a:t>this</a:t></a:r><a:r><a:rPr lang="en-US" dirty="0"/><a:t> and </a:t></a:r><a:r><a:rPr lang="en-US" dirty="0"><a:latin typeface="Mono"/></a:rPr><a:t>that</a:t></a:r></a:p></p:txBody>',
		);
	});

	it("keeps notes' line breaks and adds a blank paragraph where the Markdown has blank lines", async () => {
		expect(await notesXml("One *two*\nthree\n\nFour ~~five~~")).toContain(
			'<a:p><a:r><a:rPr lang="en-US" dirty="0"/><a:t>One </a:t></a:r><a:r><a:rPr lang="en-US" i="1" dirty="0"/><a:t>two</a:t></a:r><a:br><a:rPr lang="en-US" dirty="0"/></a:br><a:r><a:rPr lang="en-US" dirty="0"/><a:t>three</a:t></a:r></a:p><a:p><a:endParaRPr lang="en-US" dirty="0"/></a:p><a:p><a:r><a:rPr lang="en-US" dirty="0"/><a:t>Four </a:t></a:r><a:r><a:rPr lang="en-US" strike="sngStrike" dirty="0"/><a:t>five</a:t></a:r></a:p></p:txBody>',
		);
	});

	it("links notes' text when the Markdown has a link", async () => {
		const zip = await writeZip(
			deckOf([
				{
					notes: "See [Slidev](https://sli.dev) and [again](https://sli.dev)",
					states: [{ clicks: 0, items: [] }],
				},
			]),
		);

		expect({
			notes: (await part(zip, "ppt/notesSlides/notesSlide1.xml"))?.match(
				/<a:hlinkClick r:id="rId\d+"\/>/g,
			),
			rels: (
				await part(zip, "ppt/notesSlides/_rels/notesSlide1.xml.rels")
			)?.match(/<Relationship Id="rId\d+" Type="[^"]*hyperlink"[^>]*>/g),
		}).toEqual({
			notes: ['<a:hlinkClick r:id="rId3"/>', '<a:hlinkClick r:id="rId3"/>'],
			rels: [
				'<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://sli.dev" TargetMode="External"/>',
			],
		});
	});

	it("bullets and numbers notes' list items when the Markdown has lists", async () => {
		expect(
			(await notesXml("- one\n  1. two\n  2. three\n- four"))?.match(
				/<a:pPr.*?<\/a:pPr>/g,
			),
		).toEqual([
			'<a:pPr marL="285750" indent="-285750"><a:buChar char="•"/></a:pPr>',
			'<a:pPr marL="571500" indent="-285750"><a:buAutoNum type="arabicPeriod" startAt="1"/></a:pPr>',
			'<a:pPr marL="571500" indent="-285750"><a:buAutoNum type="arabicPeriod" startAt="1"/></a:pPr>',
			'<a:pPr marL="285750" indent="-285750"><a:buChar char="•"/></a:pPr>',
		]);
	});

	it("bolds notes' headings and breaks their lines where the Markdown breaks them", async () => {
		expect(await notesXml("# Intro\nOne  \nTwo")).toContain(
			'<a:p><a:r><a:rPr lang="en-US" b="1" dirty="0"/><a:t>Intro</a:t></a:r></a:p><a:p><a:r><a:rPr lang="en-US" dirty="0"/><a:t>One</a:t></a:r><a:br><a:rPr lang="en-US" dirty="0"/></a:br><a:r><a:rPr lang="en-US" dirty="0"/><a:t>Two</a:t></a:r></a:p></p:txBody>',
		);
	});

	it("writes notes' code blocks in the monospace font, a paragraph per line", async () => {
		expect(await notesXml("```\nconst a = 1;\nconst b = 2;\n```")).toContain(
			'<a:p><a:r><a:rPr lang="en-US" dirty="0"><a:latin typeface="Mono"/></a:rPr><a:t>const a = 1;</a:t></a:r></a:p><a:p><a:r><a:rPr lang="en-US" dirty="0"><a:latin typeface="Mono"/></a:rPr><a:t>const b = 2;</a:t></a:r></a:p></p:txBody>',
		);
	});

	it("writes one slide per click state when withClicks is set", async () => {
		const zip = await writeZip(
			deckOf([
				{
					states: [
						{ clicks: 0, items: [] },
						{ clicks: 1, items: [textItem([[run("Shown")]])] },
					],
				},
			]),
			true,
		);

		expect({
			first: await part(zip, "ppt/slides/slide1.xml"),
			second: await part(zip, "ppt/slides/slide2.xml"),
		}).toEqual({
			first: expect.not.stringContaining("Shown"),
			second: expect.stringContaining("<a:t>Shown</a:t>"),
		});
	});

	it("numbers each click slide's note by its Slidev click when withClicks is set", async () => {
		const zip = await writeZip(
			deckOf([
				{
					states: [
						{ clicks: 2, items: [] },
						{ clicks: 3, items: [] },
					],
				},
			]),
			true,
		);

		expect({
			first: await part(zip, "ppt/notesSlides/notesSlide1.xml"),
			second: await part(zip, "ppt/notesSlides/notesSlide2.xml"),
		}).toEqual({
			first: expect.stringContaining("<a:t>Slide 1, click 2</a:t>"),
			second: expect.stringContaining("<a:t>Slide 1, click 3</a:t>"),
		});
	});

	it("counts a slide's title once when it shows on several click slides", async () => {
		const title = (size: number) =>
			textItem([[run("Title", { size })]], { heading: true });
		const zip = await writeZip(
			deckOf([
				{
					states: [0, 1, 2].map((clicks) => ({
						clicks,
						items: [title(40)],
					})),
				},
				{ states: [{ clicks: 0, items: [title(24)] }] },
				{ states: [{ clicks: 0, items: [title(24)] }] },
			]),
			true,
		);

		expect(
			firstLevel(
				await part(zip, "ppt/slideMasters/slideMaster1.xml"),
				"p:titleStyle",
			)?.properties,
		).toContain('<a:defRPr sz="2400"');
	});

	it("plays a slide's transition only into its first click state when withClicks is set", async () => {
		const zip = await writeZip(
			deckOf([
				{
					states: [
						{ clicks: 0, items: [] },
						{ clicks: 1, items: [] },
					],
					transition: "slide-left",
				},
			]),
			true,
		);

		expect({
			first: await part(zip, "ppt/slides/slide1.xml"),
			second: await part(zip, "ppt/slides/slide2.xml"),
		}).toEqual({
			first: expect.stringContaining('<p:push dir="l"/>'),
			second: expect.not.stringContaining("<p:transition"),
		});
	});

	it("fades shapes in when a click shows them", async () => {
		const zip = await writeZip(
			deckOf([
				{
					states: [
						{ clicks: 0, items: [] },
						{ clicks: 1, items: [textItem([[run("Shown")]])] },
					],
				},
			]),
		);

		expect(await part(zip, "ppt/slides/slide1.xml")).toContain(
			'<p:par><p:cTn id="3" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" grpId="0" nodeType="clickEffect">',
		);
	});

	it("fades a list's paragraph out when a click hides it", async () => {
		const shown = listItem([listParagraph("A"), listParagraph("B")]);
		const hidden = listItem([
			listParagraph("A"),
			listParagraph("B", { visible: false }),
		]);
		const zip = await writeZip(
			deckOf([
				{
					states: [
						{ clicks: 0, items: [shown] },
						{ clicks: 1, items: [hidden] },
					],
				},
			]),
		);

		expect(await part(zip, "ppt/slides/slide1.xml")).toMatch(
			/presetClass="exit".*<p:pRg st="1" end="1"\/>.*<p:bldP spid="2" grpId="0" build="p"\/>/s,
		);
	});
});

describe(dashPattern, () => {
	it.each([
		// Two widths long, with one-width gaps.
		["dashed", 6, undefined, { gap: 100000, length: 200000 }],
		// Fourteen dashes fit 244.1px with gaps of 5.85px, closer to 6px than thirteen's 7.34px.
		["dashed", 6, 244.1, { gap: 97564, length: 200000 }],
		// Drawn 2px wide, three widths long with two-width gaps: 96 dashes fit 960px with gaps of 4.04px.
		["dashed", 2.2, 960, { gap: 183732, length: 272727 }],
		["dotted", 4, undefined, { gap: 100000, length: 100000 }],
	] as const)(
		"draws a %s line %spx wide and %spx long as Blink does",
		(style, width, length, pattern) => {
			expect(dashPattern(style, width, length)).toEqual(pattern);
		},
	);
});

describe(horizontalPlacement, () => {
	const box = { height: 40, width: 200, x: 100, y: 0 };

	it("grows one-line text by its slack when its container has room", () => {
		expect(horizontalPlacement(box, { ...style, nowrap: false }, 0)).toEqual({
			left: 100,
			width: 212,
		});
	});

	it("grows centered text evenly within the room its container leaves", () => {
		expect(
			horizontalPlacement(
				box,
				{
					...style,
					align: "center",
					available: { left: 0, right: 304 },
					nowrap: false,
				},
				0,
			),
		).toEqual({ left: 96, width: 208 });
	});

	it("grows right-aligned text leftward only as far as its container's left edge", () => {
		expect(
			horizontalPlacement(
				box,
				{
					...style,
					align: "right",
					available: { left: 95, right: 300 },
					nowrap: false,
				},
				0,
			),
		).toEqual({ left: 95, width: 205 });
	});

	it("grows wrapped text only by its rounding slack when it wraps", () => {
		expect(
			horizontalPlacement(
				box,
				{
					...style,
					available: { left: 0, right: 300 },
					nowrap: false,
					wrapWidth: 150,
				},
				0,
			),
		).toEqual({ left: 100, width: 152 });
	});

	it.each([
		// Half way to where the next word would fit.
		[153, 151.5],
		// No more than its rounding slack.
		[170, 152],
	])(
		"grows wrapped text no closer than half way to keeping a word the browser wrapped, %spx wide with it",
		(wrapLimit, width) => {
			expect(
				horizontalPlacement(
					box,
					{
						...style,
						available: { left: 0, right: 300 },
						nowrap: false,
						wrapLimit,
						wrapWidth: 150,
					},
					0,
				),
			).toEqual({ left: 100, width });
		},
	);

	it("leaves text as wide as the browser drew it when it can't wrap", () => {
		expect(horizontalPlacement(box, { ...style, nowrap: true }, 0)).toEqual({
			left: 100,
			width: 200,
		});
	});
});

describe(markerFor, () => {
	it.each([
		["disc", { code: "25CF", kind: "keyword" }],
		['"→"', { code: "2192", kind: "string" }],
		['"🙂"', { code: "2022", kind: "string" }],
		['""', undefined],
		["upper-roman", { kind: "number", style: "romanUcPeriod" }],
		["cjk-decimal", { kind: "number", style: "arabicPeriod" }],
		["disclosure-closed", { code: "25B8", kind: "string" }],
		["none", undefined],
	])("writes %s as %j", (listStyle, marker) => {
		expect(markerFor(listStyle)).toEqual(marker);
	});
});

describe(numberLength, () => {
	it.each([
		[12, "decimal", 2],
		[8, "lower-roman", 4],
		[27, "upper-alpha", 2],
	])("counts %d in %s as %d characters", (number, listStyle, length) => {
		expect(numberLength(number, listStyle)).toBe(length);
	});
});

describe(codePitch, () => {
	const code = (baselines: number[]) => ({
		baselines,
		box: { height: 100, width: 400, x: 0, y: 0 },
		decorative: false,
		key: "code",
		kind: "code" as const,
		paragraphs: [],
		style: { ...style, lineHeight: 18, nowrap: true },
	});

	it("measures the pitch from the baselines when it differs from the line height", () => {
		expect(codePitch(code([10, 30, 50, 90]))).toBe(20);
	});

	it("keeps the line height when the baselines match it", () => {
		expect(codePitch(code([10, 28.1, 46]))).toBe(18);
	});
});

describe(timeline, () => {
	const text = textItem([[run("Words")]]);

	it("merges an item across click states when it stays in place", () => {
		expect(
			timeline({
				notes: "",
				number: 1,
				states: [
					{ clicks: 0, items: [text] },
					{ clicks: 1, items: [text] },
				],
			}),
		).toEqual([
			expect.objectContaining({ item: text, states: new Set([0, 1]) }),
		]);
	});

	it("orders an item that first shows on a later click beneath what it's beneath in that click", () => {
		const box = boxItem();

		expect(
			timeline({
				notes: "",
				number: 1,
				states: [
					{ clicks: 0, items: [text] },
					{ clicks: 1, items: [box, text] },
				],
			}),
		).toEqual([
			expect.objectContaining({ item: box }),
			expect.objectContaining({ item: text }),
		]);
	});
});
