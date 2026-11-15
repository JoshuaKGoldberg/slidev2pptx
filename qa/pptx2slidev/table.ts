import type { Placement } from "./geometry.ts";
import type { SlideContext } from "./shapes.ts";
import type { RunStyle } from "./text.ts";
import type { Theme } from "./theme.ts";

import { EMU_PER_PX } from "./geometry.ts";
import { px, styleOf } from "./html.ts";
import { convertText } from "./text.ts";
import { childNamed, childrenNamed, descendantNamed } from "./xml.ts";

/** A visible border line, or none where a part or cell turns one off. */
type Border =
	| "none"
	| {
			color: string;
			pattern: "dashed" | "dotted" | "double" | "solid";
			width: number;
	  };

/** A cell, by the grid rows and columns it spans (first and last, inclusive), with its formatting. */
interface Cell {
	columns: [number, number];
	element: Element;
	format: CellFormat;
	rows: [number, number];
}

interface CellFormat {
	borders: Partial<Record<Side, Border>>;
	fill?: string;
	text: Partial<Pick<RunStyle, "bold" | "color" | "font" | "italic">>;
}

/** Rows and columns a table style part covers, first and last, inclusive. */
interface Region {
	columns: readonly [number, number];
	rows: readonly [number, number];
}

type Side = "bottom" | "left" | "right" | "top";

/** Which of a part's borders each side of a cell takes: its outer one on the part's edge, else the inside one. */
const SIDES: { inside: string; outer: string; side: Side; tag: string }[] = [
	{ inside: "a:insideV", outer: "a:left", side: "left", tag: "a:lnL" },
	{ inside: "a:insideV", outer: "a:right", side: "right", tag: "a:lnR" },
	{ inside: "a:insideH", outer: "a:top", side: "top", tag: "a:lnT" },
	{ inside: "a:insideH", outer: "a:bottom", side: "bottom", tag: "a:lnB" },
];

const ANCHORS: Partial<Record<string, string>> = {
	b: "bottom",
	ctr: "middle",
};

/**
 * Converts a table graphic frame into an HTML table on its grid, styled as PowerPoint draws it:
 * its table style's parts (whole table, bands, first and last rows and columns, corners) in order,
 * then each cell's own fill, borders, margins, and text.
 * pptx-svg draws tables without their styles, so this reads the table's XML directly.
 */
export function convertTable(
	context: SlideContext,
	id: string,
	placement: Placement,
) {
	const table = descendantNamed(context.raw.get(id), "a:tbl");
	if (!table) {
		return "";
	}

	const properties = childNamed(table, "a:tblPr");
	const flags = new Set(
		["bandCol", "bandRow", "firstCol", "firstRow", "lastCol", "lastRow"].filter(
			(name) => ["1", "true"].includes(properties?.getAttribute(name) ?? ""),
		),
	);
	const styleId = childNamed(properties, "a:tableStyleId")?.textContent;
	const style = styleId ? context.theme.tableStyle(styleId) : undefined;
	if (styleId && !style) {
		context.lose(`Drew table ${id} without its style ${styleId}`);
	}

	const background = [
		...(properties?.children ?? []),
		...(childNamed(style, "a:tblBg")?.children ?? []),
	].find(
		(child) =>
			/^a:(?:blip|grad|patt|solid)Fill$|^a:fillRef$/.test(child.tagName) ||
			(child.tagName === "a:effectLst" && child.children.length),
	);
	if (background) {
		context.lose(`Dropped table ${id}'s ${background.tagName.slice(2)}`);
	}

	const columns = childrenNamed(
		childNamed(table, "a:tblGrid"),
		"a:gridCol",
	).map((column) => Number(column.getAttribute("w") ?? 0) / EMU_PER_PX);
	const rows = childrenNamed(table, "a:tr");
	const size = { columns: columns.length, rows: rows.length };

	// Each grid column has a cell in each row; merged-over ones say so, or sit under a row span.
	const cells: Cell[] = [];
	const covered = new Set<string>();
	for (const [row, element] of rows.entries()) {
		for (const [column, cell] of childrenNamed(element, "a:tc").entries()) {
			if (
				cell.getAttribute("hMerge") === "1" ||
				cell.getAttribute("vMerge") === "1" ||
				covered.has(`${String(row)},${String(column)}`)
			) {
				continue;
			}

			const span = (name: string) =>
				Math.max(1, Number(cell.getAttribute(name) ?? 1)) - 1;
			const extent = {
				columns: [
					column,
					Math.min(size.columns - 1, column + span("gridSpan")),
				] as [number, number],
				rows: [row, Math.min(size.rows - 1, row + span("rowSpan"))] as [
					number,
					number,
				],
			};
			for (let r = extent.rows[0]; r <= extent.rows[1]; r += 1) {
				for (let c = extent.columns[0]; c <= extent.columns[1]; c += 1) {
					covered.add(`${String(r)},${String(c)}`);
				}
			}

			cells.push({
				...extent,
				element: cell,
				format: formatOf(
					context,
					style,
					partsOf(flags, size, extent),
					extent,
					cell,
					id,
				),
			});
		}
	}

	const edges = edgesOf(cells);
	const outer = (pattern: RegExp) =>
		Math.max(
			0,
			...[...edges]
				.filter(([key]) => pattern.test(key))
				.map(([, border]) => widthOf(border)),
		) / 2;
	const left = outer(/^v\d+,0$/);
	const right = outer(new RegExp(`^v\\d+,${String(size.columns)}$`));
	const top = outer(/^h0,/);
	const html = rows.map((row, index) => {
		const height = Number(row.getAttribute("h") ?? 0) / EMU_PER_PX;
		return `<tr${height ? ` style="height:${px(height)}"` : ""}>${cells
			.filter((cell) => cell.rows[0] === index)
			.map((cell) => convertCell(context, cell, edges))
			.join("")}</tr>`;
	});

	// A browser puts half of a table's outer borders inside its box, around the grid.
	return `<table class="pptx-shape pptx-table" style="${styleOf({
		left: px(placement.x - left),
		top: px(placement.y - top),
		width: px(columns.reduce((sum, value) => sum + value, left + right)),
	})}"><colgroup>${columns
		.map((value) => `<col style="width:${px(value)}">`)
		.join("")}</colgroup><tbody>${html.join("")}</tbody></table>`;
}

/** A line (a:ln, or a cell's a:lnL and the like) as a border, over the border it overrides. */
function borderOf(
	theme: Theme,
	line: Element | undefined,
	inherited: Border | undefined,
): Border | undefined {
	if (!line) {
		return inherited;
	}

	if (childNamed(line, "a:noFill")) {
		return "none";
	}

	const color =
		theme.color(childNamed(line, "a:solidFill")) ??
		(inherited === "none" ? undefined : inherited?.color);
	if (!color) {
		return inherited;
	}

	const dash = childNamed(line, "a:prstDash")?.getAttribute("val") ?? "solid";
	return {
		color,
		pattern:
			line.getAttribute("cmpd") === "dbl"
				? "double"
				: dash === "solid"
					? "solid"
					: /dot/i.test(dash) && !/dash/i.test(dash)
						? "dotted"
						: "dashed",
		width: Number(line.getAttribute("w") ?? 12_700) / EMU_PER_PX,
	};
}

function convertCell(
	context: SlideContext,
	cell: Cell,
	edges: Map<string, Border>,
) {
	const properties = childNamed(cell.element, "a:tcPr");
	if (
		properties?.getAttribute("vert") &&
		properties.getAttribute("vert") !== "horz"
	) {
		context.lose("Drew a table cell's vertical text horizontally");
	}

	const body = childNamed(cell.element, "a:txBody");
	const text = convertText(
		context,
		body,
		undefined,
		undefined,
		0,
		0,
		0,
		(run, paragraph) => runStyleOf(context, body, cell.format, run, paragraph),
	);
	// Each line between two cells is the upper or left one's, so renderers without collapsed borders draw it once.
	const border = (side: Side) => {
		const edge = edges.get(edgeKey(cell, side));
		const owned =
			side === "bottom" ||
			side === "right" ||
			(side === "top" ? cell.rows[0] : cell.columns[0]) === 0;
		return owned && edge && edge !== "none"
			? `${px(edge.width)} ${edge.pattern} ${edge.color}`
			: undefined;
	};

	// PowerPoint insets text from the grid, where a browser insets it from half of a collapsed border.
	const padding = (side: Side, name: string, fallback: number) =>
		px(
			Math.max(
				0,
				Number(properties?.getAttribute(name) ?? fallback) / EMU_PER_PX -
					widthOf(edges.get(edgeKey(cell, side))) / 2,
			),
		);
	const spans = [
		cell.columns[1] > cell.columns[0]
			? ` colspan="${String(cell.columns[1] - cell.columns[0] + 1)}"`
			: "",
		cell.rows[1] > cell.rows[0]
			? ` rowspan="${String(cell.rows[1] - cell.rows[0] + 1)}"`
			: "",
	].join("");
	return `<td${spans} style="${styleOf({
		"background-color": cell.format.fill,
		"border-bottom": border("bottom"),
		"border-left": border("left"),
		"border-right": border("right"),
		"border-top": border("top"),
		"padding-bottom": padding("bottom", "marB", 45_720),
		"padding-left": padding("left", "marL", 91_440),
		"padding-right": padding("right", "marR", 91_440),
		"padding-top": padding("top", "marT", 45_720),
		"vertical-align":
			ANCHORS[properties?.getAttribute("anchor") ?? ""] ?? "top",
	})}">${text ?? ""}</td>`;
}

/** The grid line segment along a cell's side: the one at its first row or column, for spanning cells. */
function edgeKey(cell: Pick<Cell, "columns" | "rows">, side: Side, offset = 0) {
	switch (side) {
		case "bottom": {
			return `h${String(cell.rows[1] + 1)},${String(cell.columns[0] + offset)}`;
		}

		case "left": {
			return `v${String(cell.rows[0] + offset)},${String(cell.columns[0])}`;
		}

		case "right": {
			return `v${String(cell.rows[0] + offset)},${String(cell.columns[1] + 1)}`;
		}

		default: {
			return `h${String(cell.rows[0])},${String(cell.columns[0] + offset)}`;
		}
	}
}

/**
 * The border along each grid line segment, where neighboring cells can disagree:
 * a visible border wins over none, and a wider one over a narrower one, else the cell above or to the left.
 */
function edgesOf(cells: Cell[]) {
	const edges = new Map<string, Border>();
	for (const cell of cells) {
		for (const side of ["top", "right", "bottom", "left"] as const) {
			const border = cell.format.borders[side];
			const across = side === "top" || side === "bottom";
			const length = across
				? cell.columns[1] - cell.columns[0]
				: cell.rows[1] - cell.rows[0];
			for (let offset = 0; offset <= length; offset += 1) {
				const key = edgeKey(cell, side, offset);
				const existing = edges.get(key);
				if (border && widthOf(border) > widthOf(existing)) {
					edges.set(key, border);
				} else if (border && !existing) {
					edges.set(key, border);
				}
			}
		}
	}

	return edges;
}

/** A fill (a:solidFill, a:noFill, or another) among an element's children as a CSS color, over what it overrides. */
function fillOf(
	context: SlideContext,
	parent: Element | undefined,
	inherited: string | undefined,
	id: string,
) {
	for (const child of parent?.children ?? []) {
		if (child.tagName === "a:noFill") {
			return undefined;
		}

		if (child.tagName === "a:solidFill" || child.tagName === "a:fillRef") {
			return context.theme.color(child) ?? inherited;
		}

		if (/^a:(?:blip|grad|grp|patt)Fill$/.test(child.tagName)) {
			context.lose(
				`Drew a cell of table ${id} without its ${child.tagName.slice(2)}`,
			);
			return inherited;
		}
	}

	return inherited;
}

/** A cell's fill, borders, and text formatting from its table style's parts, then its own properties. */
function formatOf(
	context: SlideContext,
	style: Element | undefined,
	parts: { name: string; region: Region }[],
	cell: Region,
	element: Element,
	id: string,
): CellFormat {
	const format: CellFormat = { borders: {}, text: {} };
	const { theme } = context;
	for (const { name, region } of parts) {
		const part = childNamed(style, name);
		if (!part) {
			continue;
		}

		const cellStyle = childNamed(part, "a:tcStyle");
		format.fill = fillOf(
			context,
			childNamed(cellStyle, "a:fill") ?? cellStyle,
			format.fill,
			id,
		);
		const borders = childNamed(cellStyle, "a:tcBdr");
		const onEdge: Record<Side, boolean> = {
			bottom: cell.rows[1] === region.rows[1],
			left: cell.columns[0] === region.columns[0],
			right: cell.columns[1] === region.columns[1],
			top: cell.rows[0] === region.rows[0],
		};
		for (const { inside, outer, side } of SIDES) {
			const border = childNamed(borders, onEdge[side] ? outer : inside);
			const reference = childNamed(border, "a:lnRef");
			const color = reference && theme.color(reference);
			format.borders[side] = color
				? { color, pattern: "solid", width: 1 }
				: borderOf(theme, childNamed(border, "a:ln"), format.borders[side]);
		}

		const text = childNamed(part, "a:tcTxStyle");
		if (text) {
			const toggle = (value: null | string, current: boolean | undefined) =>
				value === "on" ? true : value === "off" ? false : current;
			format.text.bold = toggle(text.getAttribute("b"), format.text.bold);
			format.text.italic = toggle(text.getAttribute("i"), format.text.italic);
			const fontReference = childNamed(text, "a:fontRef");
			const typeface =
				childNamed(childNamed(text, "a:font"), "a:latin")?.getAttribute(
					"typeface",
				) ??
				(fontReference
					? fontReference.getAttribute("idx") === "major"
						? "+mj-lt"
						: "+mn-lt"
					: undefined);
			format.text.font = typeface ? theme.font(typeface) : format.text.font;
			format.text.color =
				theme.color(text) ?? theme.color(fontReference) ?? format.text.color;
		}
	}

	const properties = childNamed(element, "a:tcPr");
	format.fill = fillOf(context, properties, format.fill, id);
	for (const { side, tag } of SIDES) {
		format.borders[side] = borderOf(
			theme,
			childNamed(properties, tag),
			format.borders[side],
		);
	}

	if (
		[...(properties?.children ?? [])].some(
			(line) =>
				(line.tagName === "a:lnTlToBr" || line.tagName === "a:lnBlToTr") &&
				!childNamed(line, "a:noFill"),
		)
	) {
		context.lose(`Dropped a diagonal border in table ${id}`);
	}

	return format;
}

/** A cell's table style parts, in the order they apply, with the rows and columns each covers. */
function partsOf(
	flags: Set<string>,
	size: { columns: number; rows: number },
	{ columns: [column, columnEnd], rows: [row, rowEnd] }: Region,
) {
	const last = { column: size.columns - 1, row: size.rows - 1 };
	const whole: Region = { columns: [0, last.column], rows: [0, last.row] };
	const firstColumn = flags.has("firstCol") && column === 0;
	const lastColumn = flags.has("lastCol") && columnEnd === last.column;
	const firstRow = flags.has("firstRow") && row === 0;
	const lastRow = flags.has("lastRow") && rowEnd === last.row;
	const parts: { name: string; region: Region }[] = [
		{ name: "a:wholeTbl", region: whole },
	];

	// Bands count from the first row or column after a header.
	const bandColumn = column - (flags.has("firstCol") ? 1 : 0);
	if (flags.has("bandCol") && bandColumn >= 0 && !lastColumn) {
		parts.push({
			name: bandColumn % 2 ? "a:band2V" : "a:band1V",
			region: { columns: [column, columnEnd], rows: whole.rows },
		});
	}

	const bandRow = row - (flags.has("firstRow") ? 1 : 0);
	if (flags.has("bandRow") && bandRow >= 0 && !lastRow) {
		parts.push({
			name: bandRow % 2 ? "a:band2H" : "a:band1H",
			region: { columns: whole.columns, rows: [row, rowEnd] },
		});
	}

	const corner: Region = { columns: [column, columnEnd], rows: [row, rowEnd] };
	for (const [applies, name, region] of [
		[
			lastColumn,
			"a:lastCol",
			{ columns: [last.column, last.column], rows: whole.rows },
		],
		[firstColumn, "a:firstCol", { columns: [0, 0], rows: whole.rows }],
		[
			lastRow,
			"a:lastRow",
			{ columns: whole.columns, rows: [last.row, last.row] },
		],
		[firstRow, "a:firstRow", { columns: whole.columns, rows: [0, 0] }],
		[lastRow && lastColumn, "a:seCell", corner],
		[lastRow && firstColumn, "a:swCell", corner],
		[firstRow && lastColumn, "a:neCell", corner],
		[firstRow && firstColumn, "a:nwCell", corner],
	] as const) {
		if (applies) {
			parts.push({ name, region });
		}
	}

	return parts;
}

/**
 * A run's style in a table cell, from least to most specific: the presentation's default text style,
 * the table style's text formatting, the cell's list style, then the run's own properties.
 * A hyperlink takes the theme's hyperlink color unless it says to keep its text's.
 */
function runStyleOf(
	context: SlideContext,
	body: Element | undefined,
	format: CellFormat,
	properties: Element | undefined,
	paragraph: Element,
): RunStyle {
	const level = Number(
		childNamed(paragraph, "a:pPr")?.getAttribute("lvl") ?? 0,
	);
	const style: RunStyle = {
		bold: false,
		italic: false,
		size: 18,
		strike: false,
		underline: false,
	};
	const apply = (layer: Element | undefined) => {
		if (!layer) {
			return;
		}

		const flag = (attribute: string) => {
			const value = layer.getAttribute(attribute);
			return value === null ? undefined : value === "1" || value === "true";
		};

		style.bold = flag("b") ?? style.bold;
		style.italic = flag("i") ?? style.italic;
		const size = layer.getAttribute("sz");
		style.size = size ? Number(size) / 100 : style.size;
		const underline = layer.getAttribute("u");
		style.underline = underline ? underline !== "none" : style.underline;
		const strike = layer.getAttribute("strike");
		style.strike = strike ? strike !== "noStrike" : style.strike;
		style.color =
			context.theme.color(childNamed(layer, "a:solidFill")) ?? style.color;
		const typeface = childNamed(layer, "a:latin")?.getAttribute("typeface");
		style.font = typeface ? context.theme.font(typeface) : style.font;
	};

	// Without a body, these are the presentation's defaults, as for text outside placeholders.
	for (const defaults of context.inherited(undefined, level).toReversed()) {
		apply(childNamed(defaults, "a:defRPr"));
	}

	style.bold = format.text.bold ?? style.bold;
	style.italic = format.text.italic ?? style.italic;
	style.font = format.text.font ?? style.font;
	style.color = format.text.color ?? style.color;
	apply(
		childNamed(
			childNamed(
				childNamed(body, "a:lstStyle"),
				`a:lvl${String(level + 1)}pPr`,
			),
			"a:defRPr",
		),
	);
	apply(properties);

	const link = childNamed(properties, "a:hlinkClick");
	if (
		link?.getAttribute("r:id") &&
		descendantNamed(link, "ahyp:hlinkClr")?.getAttribute("val") !== "tx"
	) {
		style.color = context.theme.scheme("hlink") ?? style.color;
	}

	return style;
}

function widthOf(border: Border | undefined) {
	return border && border !== "none" ? border.width : 0;
}
