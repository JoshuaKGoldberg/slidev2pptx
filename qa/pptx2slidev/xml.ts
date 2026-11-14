import { DOMParser } from "linkedom";

export function parseXml(xml: string) {
	// linkedom's XMLDocument is a DOM Document, though its types don't say so.
	return new DOMParser().parseFromString(
		xml,
		"text/xml",
	) as unknown as Document;
}

/** An attribute's value, or undefined when it's missing or empty, as pptx-svg writes empty attributes for unset values. */
export function attributeOf(element: Element | undefined, name: string) {
	const value = element?.getAttribute(name);
	return value === null || value === "" ? undefined : value;
}

export function childNamed(element: Element | undefined, name: string) {
	return childrenNamed(element, name).at(0);
}

/** Direct children with a qualified tag name, such as "p:par". */
export function childrenNamed(element: Element | undefined, name: string) {
	return [...(element?.children ?? [])].filter(
		(child) => child.tagName === name,
	);
}

export function descendantNamed(element: Element | undefined, name: string) {
	return descendantsNamed(element, name).at(0);
}

/** Descendants with a qualified tag name, in document order. */
export function descendantsNamed(element: Element | undefined, name: string) {
	const found: Element[] = [];
	const visit = (parent: Element) => {
		for (const child of parent.children) {
			if (child.tagName === name) {
				found.push(child);
			}

			visit(child);
		}
	};

	if (element) {
		visit(element);
	}

	return found;
}

/** Alternate content's mc:Choice branch, so newer features (like Morph) win over their fallbacks. */
export function preferChoices(element: Element) {
	for (const fallback of descendantsNamed(element, "mc:Fallback")) {
		fallback.remove();
	}

	return element;
}
