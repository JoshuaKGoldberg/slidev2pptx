import * as fontkit from "fontkit";
import { execFileSync } from "node:child_process";
import path from "node:path";

import { metricsOf } from "../../src/fonts.ts";

/** A font file, as fontconfig reports it. */
interface FontFile {
	families: string[];
	file: string;
	fullNames: string[];
	italic: boolean;
	styles: string[];
	/** CSS weight, or a [min, max] range for variable fonts. */
	weight: [number, number] | number;
}

/** A font file to declare with an `@font-face` rule. */
export interface FontFace {
	family: string;
	file: string;
	italic: boolean;
	weight: [number, number] | number;
}

/** What a PowerPoint font name becomes in CSS. */
export interface ResolvedFont {
	family: string;
	italic?: boolean;
	weight?: number;
}

const POWERPOINT_FONTS =
	"/Applications/Microsoft PowerPoint.app/Contents/Resources/DFonts";

const FORMAT = "%{family}\\t%{style}\\t%{weight}\\t%{fullname}\\t%{file}\\n";

const STYLE_WEIGHTS: [RegExp, number][] = [
	[/^(?:thin|hairline)$/i, 100],
	[/^(?:extra|ultra)-?light$/i, 200],
	[/^light$/i, 300],
	[/^(?:regular|normal|book|roman)$/i, 400],
	[/^medium$/i, 500],
	[/^(?:semi|demi)-?bold$/i, 600],
	[/^bold$/i, 700],
	[/^(?:extra|ultra)-?bold$/i, 800],
	[/^(?:black|heavy)$/i, 900],
];

/** fontconfig weights to CSS weights. */
const FONTCONFIG_WEIGHTS: [number, number][] = [
	[0, 100],
	[40, 200],
	[50, 300],
	[80, 400],
	[100, 500],
	[180, 600],
	[200, 700],
	[205, 800],
	[210, 900],
];

let index: FontFile[] | undefined;

/**
 * Resolves PowerPoint font names to CSS families backed by local font files,
 * so Chromium and slidev2pptx both draw with the fonts PowerPoint does.
 */
export function createFontResolver() {
	index ??= [
		...readFonts(["fc-list", ["--format", FORMAT]]),
		...readFonts(["fc-scan", ["--format", FORMAT, POWERPOINT_FONTS]]),
	];
	const fonts = index;
	const faces = new Map<string, FontFace>();
	const missing = new Set<string>();

	const opened = new Map<string, fontkit.Font | undefined>();

	/** The fontkit face a family resolves to at a weight and style, picked out of a collection by family and style. */
	function faceFor(family: string, weight: number, italic: boolean) {
		const key = `${family}\0${String(weight)}\0${String(italic)}`;
		if (opened.has(key)) {
			return opened.get(key);
		}

		const lower = family.toLowerCase();
		const distance = (font: FontFile) =>
			(font.italic === italic ? 0 : 1000) +
			Math.abs(
				(Array.isArray(font.weight)
					? Math.min(Math.max(weight, font.weight[0]), font.weight[1])
					: font.weight) - weight,
			);
		const best = fonts
			.filter((font) =>
				font.families.some((name) => name.toLowerCase() === lower),
			)
			.sort((a, b) => distance(a) - distance(b))
			.at(0);
		let face: fontkit.Font | undefined;
		try {
			const file = best ? fontkit.openSync(best.file) : undefined;
			face =
				file && "fonts" in file
					? (file.fonts.find(
							(member) =>
								member.familyName.toLowerCase() === lower &&
								best?.styles.some(
									(style) =>
										style.toLowerCase() === member.subfamilyName.toLowerCase(),
								),
						) ?? file.fonts.at(0))
					: file;
		} catch {
			face = undefined;
		}

		opened.set(key, face);
		return face;
	}

	function addFaces(family: string, files: FontFile[]) {
		for (const file of files) {
			// CSS can't pick one face out of a collection, and the system resolves collections (like Menlo) by name anyway.
			if (/\.(?:ttc|otc)$/i.test(file.file)) {
				continue;
			}

			// fontconfig lists a variable font once with its weight range and again per named instance: keep the range.
			const key = `${family}\0${file.file}`;
			if (Array.isArray(faces.get(key)?.weight)) {
				continue;
			}

			faces.set(key, {
				family,
				file: file.file,
				italic: file.italic,
				weight: file.weight,
			});
		}
	}

	return {
		/** How wide text draws in a family at a weight and style, in em, without kerning; undefined if it has no file. */
		advance(family: string, weight: number, italic: boolean, text: string) {
			const face = faceFor(family, weight, italic);
			return face && face.layout(text, []).advanceWidth / face.unitsPerEm;
		},
		faces: () => [...faces.values()],
		/** Line metrics of the face a family resolves to at a weight and style, as slidev2pptx measures them; undefined if it has no file. */
		metrics(family: string, weight: number, italic: boolean) {
			const face = faceFor(family, weight, italic);
			return face && metricsOf(face, family);
		},
		missing: () => [...missing],
		resolve(name: string): ResolvedFont {
			const lower = name.toLowerCase();
			const byFamily = fonts.filter((font) =>
				font.families.some((family) => family.toLowerCase() === lower),
			);
			if (byFamily.length) {
				addFaces(name, byFamily);
				return { family: name };
			}

			// "League Spartan SemiBold" is League Spartan at weight 600.
			for (const font of fonts) {
				for (const family of font.families) {
					if (!lower.startsWith(family.toLowerCase() + " ")) {
						continue;
					}

					const style = parseStyle(name.slice(family.length + 1));
					if (style) {
						addFaces(
							family,
							fonts.filter((other) => other.families.includes(family)),
						);
						return { family, ...style };
					}
				}
			}

			const byFullName = fonts.find((font) =>
				font.fullNames.some((fullName) => fullName.toLowerCase() === lower),
			);
			if (byFullName) {
				addFaces(name, [byFullName]);
				return { family: name };
			}

			missing.add(name);
			return { family: name };
		},
	};
}

function cssWeight(fontconfig: number) {
	return FONTCONFIG_WEIGHTS.reduce((closest, entry) =>
		Math.abs(entry[0] - fontconfig) < Math.abs(closest[0] - fontconfig)
			? entry
			: closest,
	)[1];
}

function parseStyle(style: string) {
	const words = style.split(/\s+/);
	const italic = words.some((word) => /^(?:italic|oblique)$/i.test(word));
	const rest = words
		.filter((word) => !/^(?:italic|oblique)$/i.test(word))
		.join(" ");
	const weight = rest
		? STYLE_WEIGHTS.find(([pattern]) => pattern.test(rest))?.[1]
		: 400;
	return weight === undefined ? undefined : { italic, weight };
}

function readFonts([command, args]: [string, string[]]): FontFile[] {
	let output: string;
	try {
		output = execFileSync(command, args, {
			encoding: "utf8",
			maxBuffer: 64 * 1024 * 1024,
		});
	} catch {
		return [];
	}

	return output
		.split("\n")
		.filter(Boolean)
		.flatMap((line) => {
			const [families, styles, weight, fullNames, file] = line.split("\t");
			if (
				!file ||
				!/\.(?:otf|ttf|ttc)$/i.test(file) ||
				path.basename(file).startsWith(".")
			) {
				return [];
			}

			const range = /\[(\d+) (\d+)\]/.exec(weight);
			const styleList = styles.split(",");
			return [
				{
					families: families.split(","),
					file,
					fullNames: fullNames.split(","),
					italic: styleList.some((style) => /italic|oblique/i.test(style)),
					styles: styleList,
					weight: range
						? [cssWeight(Number(range[1])), cssWeight(Number(range[2]))]
						: cssWeight(Number(weight)),
				},
			];
		});
}
