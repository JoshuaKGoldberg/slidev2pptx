/**
 * Counts which PowerPoint features the corpus's decks use, and how often,
 * into qa/reports/census.json and qa/reports/census.md.
 * @example
 * ```shell
 * node qa/census/census.ts
 * ```
 */
import JSZip from "jszip";
import fs from "node:fs/promises";
import path from "node:path";

import type { Deck } from "../corpus/types.ts";

import { deckPath, manifestPath } from "../corpus/paths.ts";
import { reportsDirectory } from "../reports/paths.ts";
import { animationName } from "./animations.ts";

interface Tally {
	count: number;
	decks: Set<string>;
	slides: number;
}

/** Features found by counting a pattern in each slide's XML, by group. */
const SLIDE_FEATURES: Record<string, Record<string, RegExp>> = {
	Animation: {
		"After-previous effect": /nodeType="afterEffect"/g,
		"Animation timing": /<p:timing>/g,
		"Motion path": /presetClass="path"/g,
		"On-click effect": /nodeType="clickEffect"/g,
		"Paragraph build": /<p:bldP [^>]*build="p"/g,
		"Trigger on shape click": /evt="onClick"/g,
		"With-previous effect": /nodeType="withEffect"/g,
	},
	"Images and media": {
		Audio: /<a:audioFile |<a:wavAudioFile /g,
		"Cropped picture": /<a:srcRect [^/]*[lrtb]="/g,
		Picture: /<p:pic>/g,
		"Picture fill on a shape": /<p:spPr>(?:(?!<\/p:spPr>).)*<a:blipFill/gs,
		"SVG image": /<asvg:svgBlip /g,
		Video: /<a:videoFile /g,
	},
	"Rich content": {
		Chart: /<c:chart /g,
		"Embedded object (OLE)": /<p:oleObj /g,
		Equation: /<m:oMath/g,
		Ink: /<p14:contentPart |<aink:/g,
		SmartArt: /<dgm:relIds /g,
		Table: /<a:tbl>/g,
	},
	Shapes: {
		"3D": /<a:scene3d>|<a:sp3d /g,
		Connector: /<p:cxnSp>/g,
		"Custom geometry": /<a:custGeom>/g,
		"Dashed line": /<a:prstDash val="(?!solid)/g,
		Flip: /<a:xfrm [^>]*flip[HV]="1"/g,
		Glow: /<a:glow /g,
		"Gradient fill": /<a:gradFill/g,
		Group: /<p:grpSp>/g,
		"Inner shadow": /<a:innerShdw /g,
		"Line arrowhead": /<a:(?:head|tail)End [^>]*type="(?!none)/g,
		"Outer shadow": /<a:outerShdw /g,
		"Pattern fill": /<a:pattFill/g,
		Reflection: /<a:reflection /g,
		Rotation: /<a:xfrm [^>]*rot="/g,
		Shape: /<p:sp>/g,
		"Soft edges": /<a:softEdge /g,
		"Text box": /<p:cNvSpPr txBox="1"/g,
		Transparency: /<a:alpha val="/g,
	},
	Slides: {
		"Custom background": /<p:bg>/g,
		"Hidden slide": /<p:sld [^>]*show="0"/g,
		Placeholder: /<p:ph[ />]/g,
		Slide: /<p:sld[ >]/g,
	},
	Text: {
		"All caps or small caps": /<a:rPr [^>]*cap="(?!none)/g,
		"Bulleted paragraph": /<a:buChar /g,
		"Character spacing": /<a:rPr [^>]*spc="(?!0")/g,
		Emoji: /[\u{1F300}-\u{1FAFF}\u{2600}-\u{27BF}]/gu,
		Hyperlink: /<a:hlinkClick /g,
		"Line spacing": /<a:lnSpc>/g,
		"Numbered paragraph": /<a:buAutoNum /g,
		"Picture bullet": /<a:buBlip>/g,
		"Resize shape to fit text": /<a:spAutoFit/g,
		"Shrink text on overflow": /<a:normAutofit/g,
		Strikethrough: /strike="(?!noStrike)/g,
		"Superscript or subscript": /<a:rPr [^>]*baseline="(?!0")/g,
		"Text columns": /<a:bodyPr [^>]*numCol="(?!1")/g,
		"Text highlight": /<a:highlight>/g,
		"Text outline": /<a:rPr(?: [^>]*[^/])?>(?:(?!<\/a:rPr>).)*<a:ln[ >]/gs,
		Underline: /<a:rPr [^>]*u="(?!none)/g,
		"Vertical text": /<a:bodyPr [^>]*vert="(?!horz)/g,
		"WordArt warp": /<a:prstTxWarp prst="(?!textNoShape)/g,
	},
	Transitions: {
		"Auto-advance after time": /<p:transition [^>]*advTm="/g,
		Morph: /<p159:morph[ />]/g,
		Transition: /<p:transition[ >]/g,
	},
};

/** Features counted by tallying values, such as which preset shapes are used. */
const SLIDE_BREAKDOWNS: Record<string, (xml: string) => string[]> = {
	"Animation effects": (xml) =>
		[...xml.matchAll(/<p:cTn [^>]*>/g)].flatMap(([element]) => {
			const presetClass = /presetClass="(\w+)"/.exec(element)?.[1];
			const presetID = /presetID="(\d+)"/.exec(element)?.[1];
			return presetClass && presetID
				? [`${presetClass} ${animationName(presetClass, Number(presetID))}`]
				: [];
		}),
	"Font typefaces": (xml) =>
		[...xml.matchAll(/<a:(?:latin|ea|cs|sym) typeface="([^"]+)"/g)].map(
			([, typeface]) => typeface,
		),
	"Preset shapes": (xml) =>
		[...xml.matchAll(/<a:prstGeom prst="(\w+)"/g)].map(([, shape]) => shape),
	"Transition durations (ms)": (xml) =>
		transitions(xml).map(
			(transition) =>
				/p14:dur="(\d+)"/.exec(transition)?.[1] ??
				`spd=${/spd="(\w+)"/.exec(transition)?.[1] ?? "fast"}`,
		),
	"Transition effects": (xml) =>
		transitions(xml).map(
			(transition) =>
				/<(?:p|p14|p15|p159):(?!transition|sndAc|stSnd|endSnd|snd)(\w+)/.exec(
					transition,
				)?.[1] ?? "none",
		),
};

const manifest = JSON.parse(await fs.readFile(manifestPath, "utf8")) as Deck[];
const decks = manifest.filter((deck) => deck.sha256);

const features = new Map<string, Tally>();
const breakdowns = new Map<string, Map<string, Tally>>();
const deckFacts: DeckFacts[] = [];

interface DeckFacts {
	embeddedFonts: string[];
	id: string;
	layouts: number;
	masters: number;
	sections: number;
	size: string;
	slides: number;
	themeFonts: (string | undefined)[];
}

for (const deck of decks) {
	const zip = await JSZip.loadAsync(await fs.readFile(deckPath(deck.id)));
	const files = Object.keys(zip.files);
	const slides = files.filter((name) =>
		/^ppt\/slides\/slide\d+\.xml$/.test(name),
	);

	for (const slide of slides) {
		const xml = withoutFallbacks(
			(await zip.file(slide)?.async("string")) ?? "",
		);
		for (const [group, patterns] of Object.entries(SLIDE_FEATURES)) {
			for (const [label, pattern] of Object.entries(patterns)) {
				tally(
					features,
					`${group}: ${label}`,
					deck.id,
					xml.match(pattern)?.length ?? 0,
				);
			}
		}

		for (const [name, values] of Object.entries(SLIDE_BREAKDOWNS)) {
			const counts = new Map<string, number>();
			for (const value of values(xml)) {
				counts.set(value, (counts.get(value) ?? 0) + 1);
			}

			const breakdown = breakdowns.get(name) ?? new Map<string, Tally>();
			breakdowns.set(name, breakdown);
			for (const [value, count] of counts) {
				tally(breakdown, value, deck.id, count);
			}
		}

		const rels =
			(await zip
				.file(slide.replace("slides/", "slides/_rels/") + ".rels")
				?.async("string")) ?? "";
		tally(
			features,
			"Slides: Speaker notes",
			deck.id,
			(await notesText(zip, rels)) ? 1 : 0,
		);
		for (const [relationship] of rels.matchAll(/<Relationship [^>]*>/g)) {
			if (relationship.includes("/image")) {
				const extension = /Target="[^"]+\.(\w+)"/.exec(relationship)?.[1];
				tally(
					features,
					`Images and media: ${extension?.toLowerCase() ?? "unknown"} file`,
					deck.id,
					1,
				);
			}
		}
	}

	const presentation =
		(await zip.file("ppt/presentation.xml")?.async("string")) ?? "";
	const theme = (await zip.file("ppt/theme/theme1.xml")?.async("string")) ?? "";
	const size = /<p:sldSz cx="(\d+)" cy="(\d+)"/.exec(presentation);
	deckFacts.push({
		embeddedFonts: [
			...presentation.matchAll(/<p:font typeface="([^"]+)"/g),
		].map(([, font]) => font),
		id: deck.id,
		layouts: files.filter((name) =>
			/^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(name),
		).length,
		masters: files.filter((name) =>
			/^ppt\/slideMasters\/slideMaster\d+\.xml$/.test(name),
		).length,
		sections: presentation.match(/<p14:section /g)?.length ?? 0,
		size: size ? `${size[1]}x${size[2]}` : "unknown",
		slides: slides.length,
		themeFonts: [
			/<a:majorFont><a:latin typeface="([^"]*)"/.exec(theme)?.[1],
			/<a:minorFont><a:latin typeface="([^"]*)"/.exec(theme)?.[1],
		],
	});
}

await fs.mkdir(reportsDirectory, { recursive: true });
await fs.writeFile(
	path.join(reportsDirectory, "census.json"),
	JSON.stringify(
		{
			breakdowns: Object.fromEntries(
				[...breakdowns].map(([name, breakdown]) => [
					name,
					Object.fromEntries([...breakdown].sort(byDecks).map(serialize)),
				]),
			),
			decks: deckFacts,
			features: Object.fromEntries([...features].sort(byDecks).map(serialize)),
		},
		null,
		"\t",
	) + "\n",
);
await fs.writeFile(path.join(reportsDirectory, "census.md"), renderMarkdown());
console.log(
	`Counted features across ${String(decks.length)} decks into ${reportsDirectory}.`,
);

function byDecks([, a]: [string, Tally], [, b]: [string, Tally]) {
	return b.decks.size - a.decks.size || b.count - a.count;
}

async function notesText(zip: JSZip, rels: string) {
	const target = /Target="\.\.\/notesSlides\/([^"]+)"/.exec(rels)?.[1];
	if (!target) {
		return "";
	}

	const xml =
		(await zip.file(`ppt/notesSlides/${target}`)?.async("string")) ?? "";
	return [...xml.matchAll(/<p:sp>(?:(?!<\/p:sp>).)*<\/p:sp>/gs)]
		.filter(([shape]) => shape.includes('type="body"'))
		.flatMap(([shape]) => [...shape.matchAll(/<a:t>([^<]*)<\/a:t>/g)])
		.map(([, text]) => text)
		.join("")
		.trim();
}

function renderMarkdown() {
	const row = ([label, { count, decks, slides }]: [string, Tally]) =>
		`| ${label} | ${String(decks.size)} | ${String(slides)} | ${String(count)} |`;
	const lines = [
		"# Corpus Feature Census",
		"",
		`Generated by \`node qa/census/census.ts\` from ${String(decks.length)} of the ${String(manifest.length)} decks in \`qa/corpus/manifest.json\`.`,
		"Counts ignore `mc:Fallback` content, so a Morph transition counts once, as Morph.",
		"",
		"## Features",
		"",
		"| Feature | Decks | Slides | Count |",
		"| ------- | ----: | -----: | ----: |",
		...[...features]
			.filter(([, { count }]) => count > 0)
			.sort(byDecks)
			.map(row),
		"",
		"Not found in any deck: " +
			[...features]
				.filter(([, { count }]) => count === 0)
				.map(([label]) => label)
				.join(", ") +
			".",
	];

	for (const [name, breakdown] of breakdowns) {
		lines.push(
			"",
			`## ${name}`,
			"",
			"| Value | Decks | Slides | Count |",
			"| ----- | ----: | -----: | ----: |",
			...[...breakdown].sort(byDecks).slice(0, 40).map(row),
		);
		if (breakdown.size > 40) {
			lines.push("", `...and ${String(breakdown.size - 40)} more.`);
		}
	}

	const sizes = new Map<string, number>();
	for (const { size } of deckFacts) {
		sizes.set(size, (sizes.get(size) ?? 0) + 1);
	}

	lines.push(
		"",
		"## Slide sizes",
		"",
		...[...sizes].map(
			([size, count]) => `- ${size} EMU: ${String(count)} decks`,
		),
		"",
	);
	return lines.join("\n");
}

function serialize([key, { count, decks, slides }]: [string, Tally]) {
	return [key, { count, decks: decks.size, slides }] as const;
}

function tally(
	tallies: Map<string, Tally>,
	key: string,
	deck: string,
	count: number,
) {
	const existing = tallies.get(key) ?? {
		count: 0,
		decks: new Set(),
		slides: 0,
	};
	tallies.set(key, existing);
	if (count) {
		existing.count += count;
		existing.decks.add(deck);
		existing.slides += 1;
	}
}

/** Each slide transition's XML. */
function transitions(xml: string) {
	return [
		...xml.matchAll(
			/<p:transition(?: [^>]*)?\/>|<p:transition(?: [^>]*)?>.*?<\/p:transition>/gs,
		),
	].map(([transition]) => transition);
}

/** XML without mc:Fallback branches, so alternate content (like Morph's fade fallback) counts once. */
function withoutFallbacks(xml: string) {
	return xml.replaceAll(/<mc:Fallback(?: [^>]*)?>.*?<\/mc:Fallback>/gs, "");
}
