import { descendantNamed, preferChoices } from "./xml.ts";

/** A Slidev transition, as frontmatter, for the slide before the one PowerPoint transitions into. */
export interface ConvertedTransition {
	/** A note for the conversion log when Slidev can't match it exactly. */
	loss?: string;
	value: string;
}

/** PowerPoint's speeds, in ms, when a transition has no explicit duration. */
const SPEEDS: Record<string, number> = { fast: 500, med: 750, slow: 1000 };

/** Converts the transition into a PowerPoint slide into Slidev frontmatter. */
export function convertTransition(
	slide: Element,
): ConvertedTransition | undefined {
	const transition = descendantNamed(preferChoices(slide), "p:transition");
	if (!transition) {
		return undefined;
	}

	const effect = [...transition.children].find(
		(child) => !/:(?:sndAc|stSnd|endSnd|snd|extLst)$/.test(child.tagName),
	);
	const duration =
		Number(transition.getAttribute("p14:dur") ?? Number.NaN) ||
		SPEEDS[transition.getAttribute("spd") ?? "fast"];
	const name = effect?.tagName.split(":")[1];
	switch (name) {
		case "fade":
			return {
				value:
					effect?.getAttribute("thruBlk") === "1"
						? `{ name: fade-out, duration: ${String(duration)} }`
						: `{ name: fade, duration: ${String(duration)} }`,
			};
		case "morph":
			return {
				loss: "Morph became a view transition, which only morphs shapes with matching names",
				value: `{ name: view-transition, duration: ${String(duration)} }`,
			};
		case "push": {
			const direction = effect?.getAttribute("dir") ?? "l";
			const slide = {
				d: "slide-down",
				l: "slide-left",
				r: "slide-right",
				u: "slide-up",
			}[direction];
			return { value: slide ?? "slide-left" };
		}
		case undefined:
			return { value: "none" };

		default:
			return {
				loss: `The ${name} transition became a fade`,
				value: `{ name: fade, duration: ${String(duration)} }`,
			};
	}
}
