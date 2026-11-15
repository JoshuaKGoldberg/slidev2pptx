import { childNamed, descendantNamed, descendantsNamed } from "./xml.ts";

/**
 * A slide's theme, for formatting pptx-svg doesn't resolve (like tables'):
 * its colors through the slide's color map, its fonts, and the deck's table styles.
 */
export interface Theme {
	/** The CSS color of an element's color child (like a:srgbClr or a:schemeClr), with its transforms applied. */
	color: (parent: Element | undefined) => string | undefined;
	/** A typeface, with theme references like +mn-lt resolved to the theme's fonts. */
	font: (typeface: string) => string;
	/** The CSS color of a scheme color, like hlink or accent1, through the color map. */
	scheme: (name: string) => string | undefined;
	/** A table style from ppt/tableStyles.xml, by its id. */
	tableStyle: (id: string) => Element | undefined;
}

const COLOR_TAGS = new Set([
	"a:hslClr",
	"a:prstClr",
	"a:schemeClr",
	"a:scrgbClr",
	"a:srgbClr",
	"a:sysClr",
]);

/** The few preset colors table styles use; others are rare enough to read as black. */
const PRESET_COLORS: Partial<Record<string, string>> = {
	black: "000000",
	white: "FFFFFF",
};

/** The default color map, which masters normally repeat. */
const DEFAULT_COLOR_MAP: Record<string, string> = {
	bg1: "lt1",
	bg2: "lt2",
	tx1: "dk1",
	tx2: "dk2",
};

/**
 * A slide's theme from its master's theme part and color map, which its layout and itself can override.
 * @param theme The master's theme part.
 * @param parts The slide, its layout, and its master, most specific first.
 * @param tableStyles The deck's ppt/tableStyles.xml.
 */
export function createTheme(
	theme: Element | undefined,
	parts: (Element | undefined)[],
	tableStyles: Element | undefined,
): Theme {
	const scheme = descendantNamed(theme, "a:clrScheme");
	const map = new Map(Object.entries(DEFAULT_COLOR_MAP));
	// The master's p:clrMap, then each more specific part's override.
	for (const part of parts.toReversed()) {
		const mapping =
			childNamed(part, "p:clrMap") ??
			descendantNamed(childNamed(part, "p:clrMapOvr"), "a:overrideClrMapping");
		for (const attribute of mapping?.attributes ?? []) {
			map.set(attribute.name, attribute.value);
		}
	}

	const fonts = descendantNamed(theme, "a:fontScheme");
	const typeface = (name: string) =>
		childNamed(childNamed(fonts, name), "a:latin")?.getAttribute("typeface") ??
		undefined;
	const styles = new Map(
		descendantsNamed(tableStyles, "a:tblStyle").map((style) => [
			style.getAttribute("styleId") ?? "",
			style,
		]),
	);

	const baseOf = (
		color: Element,
	): undefined | { alpha: number; rgb: number[] } => {
		switch (color.tagName) {
			case "a:hslClr": {
				return {
					alpha: 1,
					rgb: fromHsl(
						Number(color.getAttribute("hue") ?? 0) / 60_000 / 360,
						Number(color.getAttribute("sat") ?? 0) / 100_000,
						Number(color.getAttribute("lum") ?? 0) / 100_000,
					),
				};
			}

			case "a:prstClr": {
				return {
					alpha: 1,
					rgb: fromHex(
						PRESET_COLORS[color.getAttribute("val") ?? ""] ?? "000000",
					),
				};
			}

			case "a:schemeClr": {
				const name = color.getAttribute("val") ?? "";
				const entry = childNamed(scheme, `a:${map.get(name) ?? name}`);
				const inner = [...(entry?.children ?? [])].find((child) =>
					COLOR_TAGS.has(child.tagName),
				);
				return inner ? resolve(inner) : undefined;
			}

			case "a:scrgbClr": {
				return {
					alpha: 1,
					rgb: ["r", "g", "b"].map((channel) =>
						fromLinear(Number(color.getAttribute(channel) ?? 0) / 100_000),
					),
				};
			}

			case "a:srgbClr": {
				return { alpha: 1, rgb: fromHex(color.getAttribute("val") ?? "") };
			}

			case "a:sysClr": {
				return {
					alpha: 1,
					rgb: fromHex(
						color.getAttribute("lastClr") ??
							(color.getAttribute("val") === "window" ? "FFFFFF" : "000000"),
					),
				};
			}

			default: {
				return undefined;
			}
		}
	};

	const resolve = (color: Element) => {
		const base = baseOf(color);
		return base && transformed(base, color);
	};

	return {
		color(parent) {
			const color = [...(parent?.children ?? [])].find((child) =>
				COLOR_TAGS.has(child.tagName),
			);
			const resolved = color && resolve(color);
			return resolved && cssColor(resolved);
		},
		font(name) {
			if (name.startsWith("+mj")) {
				return typeface("a:majorFont") ?? name;
			}

			return name.startsWith("+mn") ? (typeface("a:minorFont") ?? name) : name;
		},
		scheme(name) {
			const entry = childNamed(scheme, `a:${map.get(name) ?? name}`);
			const inner = [...(entry?.children ?? [])].find((child) =>
				COLOR_TAGS.has(child.tagName),
			);
			const resolved = inner && resolve(inner);
			return resolved && cssColor(resolved);
		},
		tableStyle: (id) => styles.get(id),
	};
}

function cssColor({ alpha, rgb }: { alpha: number; rgb: number[] }) {
	const bytes = rgb.map((channel) =>
		Math.round(Math.min(1, Math.max(0, channel)) * 255),
	);
	if (alpha >= 1) {
		return `#${bytes.map((byte) => byte.toString(16).padStart(2, "0")).join("")}`;
	}

	return `rgba(${bytes.join(", ")}, ${String(Math.round(Math.max(0, alpha) * 1000) / 1000)})`;
}

function fromHex(hex: string) {
	return [0, 2, 4].map(
		(start) => (Number.parseInt(hex.slice(start, start + 2), 16) || 0) / 255,
	);
}

function fromHsl(hue: number, saturation: number, lightness: number) {
	if (!saturation) {
		return [lightness, lightness, lightness];
	}

	const q =
		lightness < 0.5
			? lightness * (1 + saturation)
			: lightness + saturation - lightness * saturation;
	const p = 2 * lightness - q;
	return [hue + 1 / 3, hue, hue - 1 / 3].map((t) => {
		const turn = t - Math.floor(t);
		if (turn < 1 / 6) {
			return p + (q - p) * 6 * turn;
		}

		if (turn < 1 / 2) {
			return q;
		}

		return turn < 2 / 3 ? p + (q - p) * (2 / 3 - turn) * 6 : p;
	});
}

function fromLinear(value: number) {
	return value <= 0.003_130_8
		? value * 12.92
		: 1.055 * value ** (1 / 2.4) - 0.055;
}

function toHsl([red, green, blue]: number[]) {
	const max = Math.max(red, green, blue);
	const min = Math.min(red, green, blue);
	const lightness = (max + min) / 2;
	if (max === min) {
		return [0, 0, lightness];
	}

	const delta = max - min;
	const saturation =
		lightness > 0.5 ? delta / (2 - max - min) : delta / (max + min);
	const hue =
		max === red
			? (green - blue) / delta + (green < blue ? 6 : 0)
			: max === green
				? (blue - red) / delta + 2
				: (red - green) / delta + 4;
	return [hue / 6, saturation, lightness];
}

function toLinear(value: number) {
	return value <= 0.040_45 ? value / 12.92 : ((value + 0.055) / 1.055) ** 2.4;
}

/**
 * A color with its transforms applied in order.
 * PowerPoint shades and tints in linear RGB, and modulates luminance and saturation in HSL.
 */
function transformed(
	{ alpha, rgb }: { alpha: number; rgb: number[] },
	color: Element,
) {
	let result = { alpha, rgb };
	for (const transform of color.children) {
		const value = Number(transform.getAttribute("val") ?? 0) / 100_000;
		const hsl = (change: (hsl: number[]) => number[]) => {
			const [hue, saturation, lightness] = change(toHsl(result.rgb));
			result = {
				...result,
				rgb: fromHsl(
					hue - Math.floor(hue),
					Math.min(1, Math.max(0, saturation)),
					Math.min(1, Math.max(0, lightness)),
				),
			};
		};

		switch (transform.tagName) {
			case "a:alpha": {
				result = { ...result, alpha: value };
				break;
			}

			case "a:alphaMod": {
				result = { ...result, alpha: result.alpha * value };
				break;
			}

			case "a:alphaOff": {
				result = { ...result, alpha: result.alpha + value };
				break;
			}

			case "a:hueMod": {
				hsl(([hue, saturation, lightness]) => [
					hue * value,
					saturation,
					lightness,
				]);
				break;
			}

			case "a:hueOff": {
				hsl(([hue, saturation, lightness]) => [
					hue + (value * 100_000) / 60_000 / 360,
					saturation,
					lightness,
				]);
				break;
			}

			case "a:lumMod": {
				hsl(([hue, saturation, lightness]) => [
					hue,
					saturation,
					lightness * value,
				]);
				break;
			}

			case "a:lumOff": {
				hsl(([hue, saturation, lightness]) => [
					hue,
					saturation,
					lightness + value,
				]);
				break;
			}

			case "a:satMod": {
				hsl(([hue, saturation, lightness]) => [
					hue,
					saturation * value,
					lightness,
				]);
				break;
			}

			case "a:satOff": {
				hsl(([hue, saturation, lightness]) => [
					hue,
					saturation + value,
					lightness,
				]);
				break;
			}

			case "a:shade": {
				result = {
					...result,
					rgb: result.rgb.map((channel) =>
						fromLinear(toLinear(channel) * value),
					),
				};
				break;
			}

			case "a:tint": {
				result = {
					...result,
					rgb: result.rgb.map((channel) =>
						fromLinear(1 - (1 - toLinear(channel)) * value),
					),
				};
				break;
			}

			default: {
				break;
			}
		}
	}

	return result;
}
