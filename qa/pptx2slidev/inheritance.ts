import type JSZip from "jszip";

import {
	attributeOf,
	childNamed,
	childrenNamed,
	descendantNamed,
	descendantsNamed,
	parseXml,
} from "./xml.ts";

/** The parts a slide's text inherits paragraph properties from. */
export interface InheritanceSources {
	layout: Element | undefined;
	/** The layout's part name in the package, like ppt/slideLayouts/slideLayout2.xml. */
	layoutPath?: string;
	master: Element | undefined;
	/** The master's part name in the package, like ppt/slideMasters/slideMaster1.xml. */
	masterPath?: string;
	presentation: Element | undefined;
}

/** Master text styles by placeholder type: titles, body text, and everything else. */
const MASTER_STYLES: Partial<Record<string, string>> = {
	body: "p:bodyStyle",
	ctrTitle: "p:titleStyle",
	obj: "p:bodyStyle",
	subTitle: "p:bodyStyle",
	title: "p:titleStyle",
};

/** Loads the layout and master a slide inherits from. */
export async function loadInheritance(
	zip: JSZip,
	slidePath: string,
	presentation: Element | undefined,
	relationshipsOf: (part: string) => Promise<Map<string, string>>,
): Promise<InheritanceSources> {
	const load = async (part: string | undefined) =>
		part
			? parseXml((await zip.file(part)?.async("string")) ?? "<empty/>")
					.documentElement
			: undefined;
	const layoutPath = [...(await relationshipsOf(slidePath)).values()].find(
		(target) => target.includes("slideLayouts/"),
	);
	const masterPath = layoutPath
		? [...(await relationshipsOf(layoutPath)).values()].find((target) =>
				target.includes("slideMasters/"),
			)
		: undefined;
	return {
		layout: await load(layoutPath),
		layoutPath,
		master: await load(masterPath),
		masterPath,
		presentation,
	};
}

/**
 * A text body's paragraph properties for a level, from most to least specific:
 * the shape's own list style, then its layout and master placeholders', then
 * the master's text style for its placeholder type (or the presentation's
 * default for text that isn't in a placeholder).
 */
export function inheritedLevels(
	sources: InheritanceSources,
	body: Element | undefined,
	level: number,
) {
	const name = `a:lvl${String(level + 1)}pPr`;
	const placeholder = descendantNamed(
		descendantNamed(body?.parentElement ?? undefined, "p:nvPr"),
		"p:ph",
	);
	const levels = [childNamed(childNamed(body, "a:lstStyle"), name)];
	if (placeholder) {
		const type = attributeOf(placeholder, "type") ?? "body";
		const index = attributeOf(placeholder, "idx");
		const layoutShape = matchingPlaceholder(sources.layout, type, index);
		const masterShape = matchingPlaceholder(sources.master, type, undefined);
		levels.push(
			levelOf(layoutShape, name),
			levelOf(masterShape, name),
			childNamed(
				descendantNamed(sources.master, MASTER_STYLES[type] ?? "p:otherStyle"),
				name,
			),
		);
	} else {
		levels.push(
			childNamed(
				descendantNamed(sources.presentation, "p:defaultTextStyle"),
				name,
			),
		);
	}

	return levels.filter((element): element is Element => !!element);
}

/** The frame a slide placeholder without one inherits: its layout placeholder's, else its master's. */
export function inheritedFrame(
	sources: InheritanceSources,
	placeholder: Element,
) {
	const type = attributeOf(placeholder, "type") ?? "body";
	return [
		matchingPlaceholder(sources.layout, type, attributeOf(placeholder, "idx")),
		matchingPlaceholder(sources.master, type, undefined),
	]
		.map((shape) => childNamed(childNamed(shape, "p:spPr"), "a:xfrm"))
		.find((frame) => frame !== undefined);
}

/** The text body properties a placeholder inherits, from its layout's placeholder, then its master's. */
export function inheritedBodyProperties(
	sources: InheritanceSources,
	body: Element | undefined,
) {
	const placeholder = descendantNamed(
		descendantNamed(body?.parentElement ?? undefined, "p:nvPr"),
		"p:ph",
	);
	if (!placeholder) {
		return [];
	}

	const type = attributeOf(placeholder, "type") ?? "body";
	return [
		matchingPlaceholder(sources.layout, type, attributeOf(placeholder, "idx")),
		matchingPlaceholder(sources.master, type, undefined),
	].flatMap((shape) => descendantNamed(shape, "a:bodyPr") ?? []);
}

function levelOf(shape: Element | undefined, name: string) {
	return childNamed(
		childNamed(descendantNamed(shape, "p:txBody"), "a:lstStyle"),
		name,
	);
}

/** A layout or master's placeholder shape a slide placeholder inherits from: by index, else by type. */
function matchingPlaceholder(
	part: Element | undefined,
	type: string,
	index: string | undefined,
) {
	const shapes = descendantsNamed(part, "p:sp").map((shape) => ({
		placeholder: descendantNamed(descendantNamed(shape, "p:nvPr"), "p:ph"),
		shape,
	}));
	const normalized = (value: string) =>
		value === "ctrTitle" ? "title" : value === "obj" ? "body" : value;
	return (
		(index
			? shapes.find(
					({ placeholder }) => attributeOf(placeholder, "idx") === index,
				)
			: undefined) ??
		shapes.find(
			({ placeholder }) =>
				placeholder &&
				normalized(attributeOf(placeholder, "type") ?? "body") ===
					normalized(type),
		)
	)?.shape;
}

/** A paragraph property: the first value its own properties or inherited levels define. */
export function propertyOf(levels: Element[], attribute: string) {
	for (const level of levels) {
		const value = level.getAttribute(attribute);
		if (value !== null) {
			return value;
		}
	}

	return undefined;
}

/** A paragraph property's child element, such as a:lnSpc, from the first level that defines it. */
export function propertyChildOf(levels: Element[], tag: string) {
	for (const level of levels) {
		const child = childNamed(level, tag);
		if (child) {
			return child;
		}
	}

	return undefined;
}

/** The level that decides a paragraph's bullet: the first to say anything about bullets. */
export function bulletLevelOf(levels: Element[]) {
	return levels.find(
		(level) =>
			childrenNamed(level, "a:buNone").length ||
			childrenNamed(level, "a:buChar").length ||
			childrenNamed(level, "a:buAutoNum").length ||
			childrenNamed(level, "a:buBlip").length,
	);
}
