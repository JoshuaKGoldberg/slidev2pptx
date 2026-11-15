import type { createFontResolver } from "./fonts.ts";
import type { Visibility } from "./timing.ts";

import { inlineExtentOf } from "../../src/fonts.ts";
import { firstBaseline, SPACED_BASELINE } from "../../src/write.ts";
import { escapeAttribute, escapeText, px, styleOf } from "./html.ts";
import { bulletLevelOf, propertyChildOf, propertyOf } from "./inheritance.ts";
import {
	attributeOf,
	childNamed,
	childrenNamed,
	descendantsNamed,
} from "./xml.ts";

export interface TextContext {
	fonts: ReturnType<typeof createFontResolver>;
	/** Records something the conversion couldn't keep. */
	lose: (loss: string) => void;
	/** Paragraph properties a text body inherits for a level, most specific first. */
	inherited: (body: Element | undefined, level: number) => Element[];
	/** Resolves a relationship id to its target, for hyperlinks. */
	resolveLink: (id: string) => string | undefined;
	/** Wraps HTML in Slidev click directives for a visibility. */
	withVisibility: (html: string, visibility: undefined | Visibility) => string;
}

interface Line {
	runs: Element[];
	y: number;
}

interface Piece {
	child: Element;
	rendered: Element | undefined;
	style: RunStyle;
	text: string;
}

interface RunStyle {
	bold: boolean;
	color?: string;
	font?: string;
	italic: boolean;
	size: number;
	strike: boolean;
	underline: boolean;
}

const ALIGNMENTS: Partial<Record<string, string>> = {
	ctr: "center",
	dist: "justify",
	just: "justify",
	l: "left",
	r: "right",
};

/** PowerPoint's single line spacing, as a multiple of the font size. */
const SINGLE_LINE = 1.2;

/** EMU per px, as text insets and indents are in EMU. */
const EMU = 12_700;

/** Converts a shape's text body into HTML paragraphs and lists, styled from pptx-svg's resolved rendering. */
export function convertText(
	context: TextContext,
	body: Element | undefined,
	rendered: Element | undefined,
	paragraphVisibility: Map<number, Visibility> | undefined,
	fontScale: number,
	/** Where pptx-svg's coordinates put the left edge of the text's content box, to measure inherited indents from. */
	contentLeft: number,
) {
	const paragraphs = childrenNamed(body, "a:p");
	if (!paragraphs.some((paragraph) => textOf(paragraph).trim())) {
		return undefined;
	}

	const renderedBullets = bulletsByParagraph(rendered, contentLeft);

	const lines = linesByParagraph(rendered);
	const blocks: string[] = [];
	let list: undefined | { items: string[]; ordered: boolean; start: number };
	let lastSize = 18;
	const numbering = new Map<number, number>();

	const flushList = () => {
		if (list) {
			const tag = list.ordered ? "ol" : "ul";
			blocks.push(
				`<${tag}${list.ordered && list.start !== 1 ? ` start="${String(list.start)}"` : ""}>${list.items.join("")}</${tag}>`,
			);
			list = undefined;
		}
	};

	for (const [index, paragraph] of paragraphs.entries()) {
		const properties = childNamed(paragraph, "a:pPr");
		const levels = [
			...(properties ? [properties] : []),
			...context.inherited(body, Number(properties?.getAttribute("lvl") ?? 0)),
		];
		const paragraphLines = lines.get(index) ?? [];
		const firstRendered = paragraphLines.at(0)?.runs.at(0);
		const align =
			ALIGNMENTS[
				propertyOf(levels, "algn") ??
					firstRendered?.parentElement?.getAttribute("data-ooxml-para-align") ??
					"l"
			];
		const runs = convertRuns(context, paragraph, paragraphLines, fontScale);
		const size = runs.size ?? sizeOfEmpty(paragraph, fontScale) ?? lastSize;
		lastSize = size;

		const previous = lines.get(index - 1)?.at(-1);
		const first = paragraphLines.at(0);
		const lineHeight = lineHeightOf(propertyChildOf(levels, "a:lnSpc"), size);
		// PowerPoint leaves out a text body's first paragraph's space before.
		const spaceBefore =
			index === 0
				? undefined
				: (spacingOf(propertyChildOf(levels, "a:spcBef"), size) ??
					(previous && first
						? Math.max(0, first.y - previous.y - lineHeight)
						: undefined));
		const spaceAfter = spacingOf(propertyChildOf(levels, "a:spcAft"), size);

		// Indents and bullets inherited from layouts and masters only show in where pptx-svg drew them.
		// Where a centered or right-aligned line starts is its alignment's doing, not a margin.
		const drawn = renderedBullets.get(index);
		const declaredMargin = propertyOf(levels, "marL");
		const declaredIndent = propertyOf(levels, "indent");
		const marginLeft =
			declaredMargin === undefined
				? align === "left" || align === "justify"
					? (drawn?.textX ?? 0)
					: 0
				: Number(declaredMargin) / EMU;
		const indent =
			declaredIndent === undefined
				? drawn?.bulletX === undefined
					? 0
					: drawn.bulletX - drawn.textX
				: Number(declaredIndent) / EMU;
		// A paragraph in its first run's font keeps the browser's line boxes from growing around a different strut.
		const paragraphFont = runs.first?.font
			? context.fonts.resolve(runs.first.font)
			: undefined;
		const shift = paragraphFont
			? baselineShift(context, {
					exact:
						childNamed(propertyChildOf(levels, "a:lnSpc"), "a:spcPts") !==
						undefined,
					family: paragraphFont.family,
					italic: paragraphFont.italic ?? runs.first?.italic ?? false,
					lineHeight,
					size,
					weight: paragraphFont.weight ?? (runs.first?.bold ? 700 : 400),
				})
			: 0;
		const style = styleOf({
			"font-family": paragraphFont ? `'${paragraphFont.family}'` : undefined,
			"font-size": px(size),
			"font-weight":
				(paragraphFont?.weight ?? (runs.first?.bold ? 700 : 400)) === 400
					? undefined
					: (paragraphFont?.weight ?? 700),
			"line-height": px(lineHeight),
			"margin-bottom": spaceAfter ? px(spaceAfter) : undefined,
			"margin-top": spaceBefore ? px(spaceBefore) : undefined,
			"text-align": align === "left" ? undefined : align,
			...(shift ? { position: "relative", top: px(shift) } : {}),
		});
		const content = runs.html || "<br>";
		const bullet =
			bulletOf(bulletLevelOf(levels)) ??
			(drawn?.character
				? { character: drawn.character, kind: "character" as const, start: 1 }
				: undefined);

		if (bullet) {
			const ordered = bullet.kind === "number";
			if (list?.ordered !== ordered) {
				flushList();
				list = { items: [], ordered, start: bullet.start };
			}

			const level = Number(properties?.getAttribute("lvl") ?? 0);
			const count = (numbering.get(level) ?? bullet.start - 1) + 1;
			numbering.set(level, count);
			// PowerPoint draws a marker in its first run's color unless the paragraph says otherwise,
			// where a browser draws a ::marker in its list item's. (Its font follows the paragraph's.)
			// Browsers end an outside marker against its item's border box, so the item starts where PowerPoint
			// hangs the marker plus the marker's width, and its padding takes its text the rest of the way.
			const markerWidth = paragraphFont
				? (context.fonts.advance(
						paragraphFont.family,
						paragraphFont.weight ?? (runs.first?.bold ? 700 : 400),
						paragraphFont.italic ?? runs.first?.italic ?? false,
						bullet.kind === "character"
							? `${bullet.character} `
							: `${ordinalLabel(count, bullet.type)}. `,
					) ?? NaN) * size
				: NaN;
			const hanging = Number.isFinite(markerWidth) && indent < 0;
			const itemStyle = styleOf({
				color: srgbOf(propertyChildOf(levels, "a:buClr")) ?? runs.first?.color,
				"list-style-type":
					bullet.kind === "character"
						? `'${escapeText(bullet.character).replaceAll("'", "\\'")} '`
						: bullet.type,
				"margin-left": hanging
					? px(marginLeft + indent + markerWidth)
					: px(marginLeft || 18 * (level + 1)),
				"padding-left":
					hanging && -indent > markerWidth
						? px(-indent - markerWidth)
						: undefined,
			});
			list.items.push(
				`<li${clickAttributes(paragraphVisibility?.get(index))} style="${style};${itemStyle}">${clickContent(content, paragraphVisibility?.get(index))}</li>`,
			);
			continue;
		}

		flushList();
		numbering.clear();
		blocks.push(
			`<p${clickAttributes(paragraphVisibility?.get(index))} style="${styleOf({
				"padding-left": marginLeft ? px(marginLeft) : undefined,
				"text-indent": indent ? px(indent) : undefined,
			})}${marginLeft || indent ? ";" : ""}${style}">${clickContent(content, paragraphVisibility?.get(index))}</p>`,
		);
	}

	flushList();
	return blocks.join("");
}

/** Bullets pptx-svg drew in their own tspans, and where each paragraph's bullet and text start. */
function bulletOf(properties: Element | undefined) {
	if (!properties || childNamed(properties, "a:buNone")) {
		return undefined;
	}

	const character = childNamed(properties, "a:buChar")?.getAttribute("char");
	if (character) {
		return { character, kind: "character" as const, start: 1 };
	}

	const number = childNamed(properties, "a:buAutoNum");
	if (number) {
		const type = number.getAttribute("type") ?? "arabicPeriod";
		return {
			kind: "number" as const,
			start: Number(number.getAttribute("startAt") ?? 1),
			type: type.startsWith("alphaLc")
				? "lower-alpha"
				: type.startsWith("alphaUc")
					? "upper-alpha"
					: type.startsWith("romanLc")
						? "lower-roman"
						: type.startsWith("romanUc")
							? "upper-roman"
							: "decimal",
		};
	}

	return undefined;
}

function bulletsByParagraph(
	rendered: Element | undefined,
	contentLeft: number,
) {
	const bullets = new Map<
		number,
		{ bulletX?: number; character?: string; textX: number }
	>();
	for (const line of rendered?.children ?? []) {
		const index = Number(
			line.getAttribute("data-ooxml-para-idx") ?? Number.NaN,
		);
		if (Number.isNaN(index)) {
			continue;
		}

		const x = Number(line.getAttribute("x") ?? 0) - contentLeft;
		const children = [...line.children];
		const marker =
			children.length &&
			children.every((child) => !child.hasAttribute("data-ooxml-run-idx")) &&
			line.hasAttribute("data-ooxml-bu-font")
				? children
						.map((child) => child.textContent)
						.join("")
						.trim()
				: undefined;
		const existing = bullets.get(index);
		if (marker) {
			bullets.set(index, {
				...existing,
				bulletX: x,
				character: marker,
				textX: existing?.textX ?? x,
			});
		} else if (!existing || existing.textX === existing.bulletX) {
			bullets.set(index, { ...existing, textX: x });
		}
	}

	return bullets;
}

function convertRuns(
	context: TextContext,
	paragraph: Element,
	lines: Line[],
	fontScale: number,
) {
	const rendered = new Map<number, Element>();
	for (const line of lines) {
		for (const run of line.runs) {
			const index = Number(run.getAttribute("data-ooxml-run-idx") ?? 0);
			if (!rendered.has(index)) {
				rendered.set(index, run);
			}
		}
	}

	const pieces: (Piece | undefined)[] = [];
	let runIndex = 0;
	let size: number | undefined;
	let first: RunStyle | undefined;
	let previous: RunStyle | undefined;
	for (const child of paragraph.children) {
		if (child.tagName === "a:br") {
			// undefined marks a line break.
			pieces.push(undefined);
			continue;
		}

		if (child.tagName !== "a:r" && child.tagName !== "a:fld") {
			continue;
		}

		const text = descendantsNamed(child, "a:t")
			.map((element) => element.textContent)
			.join("");
		const renderedRun = rendered.get(runIndex);
		runIndex += 1;
		const style = renderedRun
			? styleOfRun(renderedRun)
			: (previous ?? fallbackStyle(child, fontScale));
		if (renderedRun) {
			style.size = scaledSize(renderedRun, fontScale);
		}

		previous = style;
		size ??= style.size;
		first ??= style;
		if (text) {
			pieces.push({ child, rendered: renderedRun, style, text });
		}
	}

	// Vue drops whitespace-only text that's an element's only child, so spaces join a neighboring run.
	for (const [index, piece] of pieces.entries()) {
		if (!piece || piece.text.trim()) {
			continue;
		}

		const before = pieces[index - 1];
		const after = pieces[index + 1];
		if (before) {
			before.text += piece.text;
			pieces[index] = { ...piece, text: "" };
		} else if (after) {
			after.text = piece.text + after.text;
			pieces[index] = { ...piece, text: "" };
		}
	}

	const html = pieces
		.map((piece) =>
			piece
				? piece.text &&
					renderRun(
						context,
						piece.child,
						piece.rendered,
						piece.style,
						preserveSpaces(piece.text),
					)
				: "<br>",
		)
		.join("");
	return { first, html, size };
}

/** Click directives for a paragraph that builds in: it appears on its entrance click. */
function clickAttributes(visibility: undefined | Visibility) {
	return visibility?.enter ? ` v-click="${String(visibility.enter)}"` : "";
}

/** A paragraph's content, hidden on its exit click when it has one. */
function clickContent(content: string, visibility: undefined | Visibility) {
	return visibility?.exit === undefined
		? content
		: `<span v-click.hide="${String(visibility.exit)}">${content}</span>`;
}

function fallbackStyle(run: Element, fontScale: number): RunStyle {
	const properties = childNamed(run, "a:rPr");
	return {
		bold: properties?.getAttribute("b") === "1",
		italic: properties?.getAttribute("i") === "1",
		size:
			(Number(properties?.getAttribute("sz") ?? 1800) / 100) * (fontScale || 1),
		strike:
			/strike/i.test(properties?.getAttribute("strike") ?? "noStrike") &&
			properties?.getAttribute("strike") !== "noStrike",
		underline:
			!!properties?.getAttribute("u") &&
			properties.getAttribute("u") !== "none",
	};
}

/** A paragraph's line height from its line spacing: PowerPoint's single spacing is 1.2 times the font size. */
function lineHeightOf(spacing: Element | undefined, size: number) {
	const percent = childNamed(spacing, "a:spcPct")?.getAttribute("val");
	if (percent) {
		return size * SINGLE_LINE * (Number(percent) / 100_000);
	}

	const points = childNamed(spacing, "a:spcPts")?.getAttribute("val");
	return points ? Number(points) / 100 : size * SINGLE_LINE;
}

/** The lines pptx-svg rendered for each paragraph index, each with its run tspans. */
function linesByParagraph(rendered: Element | undefined) {
	const lines = new Map<number, Line[]>();
	for (const line of rendered?.children ?? []) {
		const index = line.getAttribute("data-ooxml-para-idx");
		if (index === null) {
			continue;
		}

		const runs = [...line.children].filter((run) =>
			run.hasAttribute("data-ooxml-run-idx"),
		);
		if (!runs.length) {
			continue;
		}

		const paragraphLines = lines.get(Number(index)) ?? [];
		paragraphLines.push({ runs, y: Number(line.getAttribute("y")) });
		lines.set(Number(index), paragraphLines);
	}

	return lines;
}

function renderRun(
	context: TextContext,
	run: Element,
	rendered: Element | undefined,
	style: RunStyle,
	text: string,
) {
	const properties = childNamed(run, "a:rPr");
	const font = style.font ? context.fonts.resolve(style.font) : undefined;
	const weight = font?.weight ?? (style.bold ? 700 : undefined);
	const baseline = Number(
		rendered?.getAttribute("data-ooxml-baseline") ??
			properties?.getAttribute("baseline") ??
			0,
	);
	const spacing = Number(
		rendered?.getAttribute("data-ooxml-char-spacing") ?? 0,
	);
	const cap =
		rendered?.getAttribute("data-ooxml-cap") ??
		properties?.getAttribute("cap") ??
		"";
	const highlight = childNamed(
		childNamed(properties, "a:highlight"),
		"a:srgbClr",
	)?.getAttribute("val");
	const outlineWidth = Number(
		rendered?.getAttribute("data-ooxml-outline-w") ?? 0,
	);
	const outlineColor = rendered?.getAttribute("data-ooxml-outline-color");
	const linkId =
		attributeOf(rendered, "data-ooxml-hlink-rid") ??
		childNamed(properties, "a:hlinkClick")?.getAttribute("r:id");
	const href = linkId ? context.resolveLink(linkId) : undefined;
	// PowerPoint underlines hyperlinks.
	const decorations = [
		style.underline || href ? "underline" : "",
		style.strike ? "line-through" : "",
	].filter(Boolean);

	const css = styleOf({
		"-webkit-text-stroke":
			outlineWidth > 0 && outlineColor
				? `${px(outlineWidth / EMU)} #${outlineColor}`
				: undefined,
		"background-color": highlight ? `#${highlight}` : undefined,
		color: style.color,
		"font-family": font ? `'${font.family.replaceAll("'", "\\'")}'` : undefined,
		"font-size": px(style.size),
		"font-style": style.italic || font?.italic ? "italic" : undefined,
		"font-variant": cap === "small" ? "small-caps" : undefined,
		"font-weight": weight && weight !== 400 ? weight : undefined,
		"letter-spacing": spacing ? px(spacing / 100) : undefined,
		"text-decoration": decorations.length ? decorations.join(" ") : undefined,
		"text-transform": cap === "all" ? "uppercase" : undefined,
		"vertical-align": baseline > 0 ? "super" : baseline < 0 ? "sub" : undefined,
	});
	const span = `<span style="${css}">${escapeText(text)}</span>`;

	return href ? `<a href="${escapeAttribute(href)}">${span}</a>` : span;
}

/** Text with its repeated spaces kept: Vue condenses runs of white space, so all but the last become no-break spaces. */
function preserveSpaces(text: string) {
	return text.replaceAll(
		/ {2,}/g,
		(spaces) => "\u00A0".repeat(spaces.length - 1) + " ",
	);
}

function scaledSize(rendered: Element, fontScale: number) {
	const size = Number(rendered.getAttribute("font-size") ?? 18);
	const specified =
		Number(rendered.getAttribute("data-ooxml-font-size") ?? 0) / 100;
	return fontScale && Math.abs(size - specified) < 0.01
		? size * fontScale
		: size;
}

function sizeOfEmpty(paragraph: Element, fontScale: number) {
	const size = childNamed(paragraph, "a:endParaRPr")?.getAttribute("sz");
	return size ? (Number(size) / 100) * (fontScale || 1) : undefined;
}

function spacingOf(spacing: Element | undefined, size: number) {
	const points = childNamed(spacing, "a:spcPts")?.getAttribute("val");
	if (points) {
		return Number(points) / 100;
	}

	const percent = childNamed(spacing, "a:spcPct")?.getAttribute("val");
	return percent ? size * SINGLE_LINE * (Number(percent) / 100_000) : undefined;
}

function srgbOf(color: Element | undefined) {
	const value = childNamed(color, "a:srgbClr")?.getAttribute("val");
	return value ? `#${value}` : undefined;
}

function styleOfRun(run: Element): RunStyle {
	const color = run.getAttribute("data-ooxml-color");
	const fill = run.getAttribute("fill");
	return {
		bold:
			run.getAttribute("data-ooxml-bold") === "1" ||
			run.getAttribute("font-weight") === "bold",
		color: color ? `#${color}` : (fill ?? undefined),
		font:
			attributeOf(run, "data-ooxml-run-font") ??
			attributeOf(run, "font-family")?.split(",")[0].trim(),
		italic:
			run.getAttribute("data-ooxml-italic") === "1" ||
			run.getAttribute("font-style") === "italic",
		size: Number(run.getAttribute("font-size") ?? 18),
		strike:
			/strike/i.test(run.getAttribute("data-ooxml-strike") ?? "") &&
			run.getAttribute("data-ooxml-strike") !== "noStrike",
		underline:
			!!run.getAttribute("data-ooxml-underline") &&
			run.getAttribute("data-ooxml-underline") !== "none",
	};
}

function textOf(paragraph: Element) {
	return descendantsNamed(paragraph, "a:t")
		.map((element) => element.textContent)
		.join("");
}

/**
 * How far below the browser's first baseline PowerPoint draws a paragraph's, in px:
 * PowerPoint puts it by its own rule for its line spacing, as slidev2pptx models it,
 * where browsers split a line's leading evenly around its font's content area.
 * Each line moves alike, as both space them by the same pitch.
 */
function baselineShift(
	context: TextContext,
	{
		exact,
		family,
		italic,
		lineHeight,
		size,
		weight,
	}: {
		exact: boolean;
		family: string;
		italic: boolean;
		lineHeight: number;
		size: number;
		weight: number;
	},
) {
	const metrics = context.fonts.metrics(family, weight, italic);
	if (!metrics) {
		return 0;
	}

	const powerPoint = firstBaseline(
		metrics,
		size,
		lineHeight,
		exact ? metrics.baseline : SPACED_BASELINE,
	);
	const { above } = inlineExtentOf(
		family,
		metrics.ascent * size,
		metrics.descent * size,
		lineHeight,
	);
	const shift = Math.round((powerPoint - above) * 100) / 100;
	return Math.abs(shift) < 0.1 ? 0 : shift;
}

/** How a browser numbers a list item for a CSS list-style-type. */
function ordinalLabel(count: number, type: string) {
	if (type.endsWith("alpha")) {
		let label = "";
		for (let rest = count; rest > 0; rest = Math.floor((rest - 1) / 26)) {
			label = String.fromCharCode(97 + ((rest - 1) % 26)) + label;
		}

		return type.startsWith("upper") ? label.toUpperCase() : label;
	}

	if (type.endsWith("roman")) {
		const numerals: [number, string][] = [
			[1000, "m"],
			[900, "cm"],
			[500, "d"],
			[400, "cd"],
			[100, "c"],
			[90, "xc"],
			[50, "l"],
			[40, "xl"],
			[10, "x"],
			[9, "ix"],
			[5, "v"],
			[4, "iv"],
			[1, "i"],
		];
		let label = "";
		let rest = count;
		for (const [value, numeral] of numerals) {
			for (; rest >= value; rest -= value) {
				label += numeral;
			}
		}

		return type.startsWith("upper") ? label.toUpperCase() : label;
	}

	return String(count);
}
