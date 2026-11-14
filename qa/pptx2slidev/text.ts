import type { createFontResolver } from "./fonts.ts";
import type { Visibility } from "./timing.ts";

import { escapeAttribute, escapeText, px, styleOf } from "./html.ts";
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
	/** Resolves a relationship id to its target, for hyperlinks. */
	resolveLink: (id: string) => string | undefined;
	/** Wraps HTML in Slidev click directives for a visibility. */
	withVisibility: (html: string, visibility: undefined | Visibility) => string;
}

interface Line {
	runs: Element[];
	y: number;
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
) {
	const paragraphs = childrenNamed(body, "a:p");
	if (!paragraphs.some((paragraph) => textOf(paragraph).trim())) {
		return undefined;
	}

	const lines = linesByParagraph(rendered);
	const ratios = [...lines.values()]
		.flatMap((paragraphLines) =>
			paragraphLines.slice(1).map((line, index) => {
				const size = styleOfRun(line.runs[0]).size;
				return (line.y - paragraphLines[index].y) / size;
			}),
		)
		.filter((ratio) => ratio > 0.5 && ratio < 3);
	const lineRatio = ratios.length
		? ratios.reduce((sum, ratio) => sum + ratio, 0) / ratios.length
		: SINGLE_LINE;

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
		const paragraphLines = lines.get(index) ?? [];
		const firstRendered = paragraphLines.at(0)?.runs.at(0);
		const align =
			ALIGNMENTS[
				properties?.getAttribute("algn") ??
					firstRendered?.parentElement?.getAttribute("data-ooxml-para-align") ??
					"l"
			];
		const runs = convertRuns(context, paragraph, paragraphLines, fontScale);
		const size = runs.size ?? sizeOfEmpty(paragraph, fontScale) ?? lastSize;
		lastSize = size;

		const previous = lines.get(index - 1)?.at(-1);
		const first = paragraphLines.at(0);
		const lineHeight = lineHeightOf(
			properties,
			size,
			lineRatio,
			paragraphLines,
		);
		const spaceBefore =
			spacingOf(childNamed(properties, "a:spcBef"), size) ??
			(previous && first
				? Math.max(0, first.y - previous.y - lineHeight)
				: undefined);
		const spaceAfter = spacingOf(childNamed(properties, "a:spcAft"), size);

		const marginLeft = Number(properties?.getAttribute("marL") ?? 0) / EMU;
		const indent = Number(properties?.getAttribute("indent") ?? 0) / EMU;
		const style = styleOf({
			"font-size": px(size),
			"line-height": px(lineHeight),
			"margin-bottom": spaceAfter ? px(spaceAfter) : undefined,
			"margin-top": spaceBefore ? px(spaceBefore) : undefined,
			"text-align": align === "left" ? undefined : align,
		});
		const content = runs.html || "<br>";
		const bullet = bulletOf(properties);

		if (bullet) {
			const ordered = bullet.kind === "number";
			if (list?.ordered !== ordered) {
				flushList();
				list = { items: [], ordered, start: bullet.start };
			}

			const level = Number(properties?.getAttribute("lvl") ?? 0);
			const count = (numbering.get(level) ?? bullet.start - 1) + 1;
			numbering.set(level, count);
			const itemStyle = styleOf({
				"list-style-type":
					bullet.kind === "character"
						? `'${escapeText(bullet.character).replaceAll("'", "\\'")} '`
						: bullet.type,
				"margin-left": px(marginLeft || 18 * (level + 1)),
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

	let html = "";
	let runIndex = 0;
	let size: number | undefined;
	let previous: RunStyle | undefined;
	for (const child of paragraph.children) {
		if (child.tagName === "a:br") {
			html += "<br>";
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
		if (!text) {
			continue;
		}

		html += renderRun(context, child, renderedRun, style, text);
	}

	return { html, size };
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

function lineHeightOf(
	properties: Element | undefined,
	size: number,
	lineRatio: number,
	lines: Line[],
) {
	const spacing = childNamed(properties, "a:lnSpc");
	const percent = childNamed(spacing, "a:spcPct")?.getAttribute("val");
	if (percent) {
		return size * SINGLE_LINE * (Number(percent) / 100_000);
	}

	const points = childNamed(spacing, "a:spcPts")?.getAttribute("val");
	if (points) {
		return Number(points) / 100;
	}

	const last = lines.at(-1);
	const pitch =
		last && lines.length > 1
			? (last.y - lines[0].y) / (lines.length - 1)
			: undefined;
	return pitch && pitch >= size * 0.8 && pitch <= size * 2.5
		? pitch
		: size * lineRatio;
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
	const decorations = [
		style.underline ? "underline" : "",
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

	const linkId =
		attributeOf(rendered, "data-ooxml-hlink-rid") ??
		childNamed(properties, "a:hlinkClick")?.getAttribute("r:id");
	const href = linkId ? context.resolveLink(linkId) : undefined;
	return href ? `<a href="${escapeAttribute(href)}">${span}</a>` : span;
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
