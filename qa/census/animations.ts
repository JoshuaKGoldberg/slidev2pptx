/** Names for common PowerPoint animation effect preset IDs, by preset class. */
const NAMES: Partial<Record<string, Partial<Record<number, string>>>> = {
	emph: {
		1: "Fill Color",
		3: "Font Color",
		6: "Grow/Shrink",
		8: "Spin",
		9: "Transparency",
		26: "Teeter",
		32: "Pulse",
	},
	entr: {
		1: "Appear",
		2: "Fly In",
		3: "Blinds",
		10: "Fade",
		12: "Peek In",
		16: "Split",
		21: "Wheel",
		22: "Wipe",
		23: "Zoom",
		26: "Bounce",
		31: "Grow & Turn",
		37: "Rise Up",
		42: "Ascend",
		47: "Descend",
		53: "Basic Zoom",
	},
	exit: {
		1: "Disappear",
		2: "Fly Out",
		10: "Fade",
		22: "Wipe",
		23: "Zoom",
		31: "Shrink & Turn",
		47: "Descend",
		53: "Basic Zoom",
	},
};

export function animationName(presetClass: string, presetID: number) {
	return NAMES[presetClass]?.[presetID] ?? `#${String(presetID)}`;
}
