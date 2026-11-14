export function escapeAttribute(value: string) {
	// Newlines would end Markdown's HTML block at a blank line, so encode them.
	return escapeText(value)
		.replaceAll('"', "&quot;")
		.replaceAll(/\r?\n/g, "&#10;");
}

/** Escapes text for Slidev slide content: HTML, plus braces so Vue doesn't read {{ }} as interpolation. */
export function escapeText(text: string) {
	return text
		.replaceAll("&", "&amp;")
		.replaceAll("<", "&lt;")
		.replaceAll(">", "&gt;")
		.replaceAll("{", "&#123;")
		.replaceAll("}", "&#125;")
		.replaceAll(/\r?\n/g, " ");
}

export function px(value: number) {
	return `${String(Math.round(value * 100) / 100)}px`;
}

/** Renders a style object as an inline style attribute value, skipping empty values. */
export function styleOf(
	declarations: Record<string, number | string | undefined>,
) {
	return (
		Object.entries(declarations)
			.filter(([, value]) => value !== undefined && value !== "")
			// Style attributes are double-quoted, so values quote with single quotes.
			.map(
				([property, value]) =>
					`${property}:${String(value).replaceAll('"', "'")}`,
			)
			.join(";")
	);
}
