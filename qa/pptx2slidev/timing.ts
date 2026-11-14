import { animationName } from "../census/animations.ts";
import {
	childNamed,
	childrenNamed,
	descendantNamed,
	descendantsNamed,
} from "./xml.ts";

export interface SlideTiming {
	/** How many clicks the slide has. */
	clicks: number;
	/** Effects that have no Slidev equivalent, so the converted slide doesn't animate them. */
	dropped: string[];
	/** Effect names by shape id, for the conversion log. */
	effects: Map<string, string[]>;
	/** Paragraph-level visibility by shape id, then paragraph index, for paragraph builds. */
	paragraphs: Map<string, Map<number, Visibility>>;
	/** Shape-level visibility by shape id. */
	shapes: Map<string, Visibility>;
}

/** When a shape (or some of its paragraphs) appears and disappears, in Slidev clicks. */
export interface Visibility {
	/** The click it disappears on, from an exit effect. */
	exit?: number;
	/** The click it appears on, from an entrance effect; 0 when it animates in as the slide starts. */
	enter?: number;
}

/** Reads a slide's main animation sequence into the clicks its shapes appear and disappear on. */
export function readTiming(slide: Element): SlideTiming {
	const timing: SlideTiming = {
		clicks: 0,
		dropped: [],
		effects: new Map(),
		paragraphs: new Map(),
		shapes: new Map(),
	};
	const root = childNamed(
		childNamed(
			childNamed(
				childNamed(descendantNamed(slide, "p:timing"), "p:tnLst"),
				"p:par",
			),
			"p:cTn",
		),
		"p:childTnLst",
	);
	if (!root) {
		return timing;
	}

	for (const sequence of childrenNamed(root, "p:seq")) {
		const container = childNamed(sequence, "p:cTn");
		if (container?.getAttribute("nodeType") !== "mainSeq") {
			timing.dropped.push(
				"an interactive sequence triggered by clicking a shape",
			);
			continue;
		}

		let click = 0;
		for (const step of childrenNamed(
			childNamed(container, "p:childTnLst"),
			"p:par",
		)) {
			const effects = descendantsNamed(step, "p:cTn").filter((node) =>
				node.hasAttribute("presetClass"),
			);
			if (effects[0]?.getAttribute("nodeType") === "clickEffect") {
				click += 1;
			}

			for (const effect of effects) {
				applyEffect(timing, effect, click);
			}
		}

		timing.clicks = Math.max(timing.clicks, click);
	}

	return timing;
}

function applyEffect(timing: SlideTiming, effect: Element, click: number) {
	const presetClass = effect.getAttribute("presetClass") ?? "";
	const name = `${presetClass} ${animationName(presetClass, Number(effect.getAttribute("presetID")))}`;
	const target = descendantNamed(effect, "p:spTgt");
	const id = target?.getAttribute("spid");
	if (!id) {
		timing.dropped.push(`${name} with no shape target`);
		return;
	}

	timing.effects.set(id, [...(timing.effects.get(id) ?? []), name]);
	if (presetClass !== "entr" && presetClass !== "exit") {
		timing.dropped.push(`${name} on shape ${id}`);
		return;
	}

	const range = descendantNamed(target, "p:pRg");
	const visibilities: Visibility[] = [];
	if (range) {
		const paragraphs =
			timing.paragraphs.get(id) ?? new Map<number, Visibility>();
		timing.paragraphs.set(id, paragraphs);
		for (
			let index = Number(range.getAttribute("st"));
			index <= Number(range.getAttribute("end"));
			index += 1
		) {
			const visibility = paragraphs.get(index) ?? {};
			paragraphs.set(index, visibility);
			visibilities.push(visibility);
		}
	} else {
		const visibility = timing.shapes.get(id) ?? {};
		timing.shapes.set(id, visibility);
		visibilities.push(visibility);
	}

	for (const visibility of visibilities) {
		if (presetClass === "entr") {
			visibility.enter ??= click;
		} else {
			visibility.exit = click;
		}
	}
}
