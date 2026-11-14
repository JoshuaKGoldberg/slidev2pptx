/**
 * Builds qa/corpus/manifest.json: one entry per unique PowerPoint deck linked
 * from a joshuakgoldberg.com speaking entry, newest first.
 * @example
 * ```shell
 * node qa/corpus/manifest.ts ~/repos/dot-com/src/content/speaking
 * ```
 */
import fs from "node:fs/promises";
import path from "node:path";
import { parse } from "yaml";

import type { Deck } from "./types.ts";

import { manifestPath } from "./paths.ts";

interface SpeakingEntry {
	category: string;
	/** The entry's file name, which orders entries on the same date. */
	file: string;
	/** YAML 1.2 has no timestamps, so dates stay ISO strings like "2023-08-08". */
	date: string;
	event: string;
	links?: Record<string, string>;
	title?: string;
}

const directory = process.argv[2];
if (!directory) {
	throw new Error("Usage: node qa/corpus/manifest.ts <speaking directory>");
}

const entries = await Promise.all(
	(await fs.readdir(directory))
		.filter((file) => file.endsWith(".md"))
		.map(async (file) => {
			const source = await fs.readFile(path.join(directory, file), "utf8");
			const frontmatter = /^---\n(.*?)\n---/s.exec(source)?.[1];
			if (!frontmatter) {
				throw new Error(`${file} has no frontmatter.`);
			}

			return { ...(parse(frontmatter) as object), file } as SpeakingEntry;
		}),
);

const decks = new Map<string, Deck>();

for (const { url, ...entry } of entries
	.flatMap(({ links, ...entry }) =>
		links?.Slides ? [{ ...entry, url: links.Slides }] : [],
	)
	.sort(
		(a, b) =>
			// Newest first; same-day entries in reverse file name order, matching earlier ad hoc downloads' numbering.
			b.date.localeCompare(a.date) || compareCodePoints(b.file, a.file),
	)) {
	const { date } = entry;
	const title = entry.title ?? entry.event;
	const existing = decks.get(url);
	if (existing) {
		existing.events.push({ date, event: entry.event });
		// Entries are newest first, so a differently titled older entry is the deck's original talk, and the newer one likely links to the wrong deck.
		if (existing.title !== title) {
			console.warn(
				`Warning: "${existing.title}" links to the same deck as the older "${title}"; using the older title.`,
			);
			existing.id = `${date}-${slugify(title)}`;
			existing.title = title;
		}

		continue;
	}

	decks.set(url, {
		category: entry.category,
		events: [{ date, event: entry.event }],
		id: `${date}-${slugify(title)}`,
		number: decks.size + 1,
		title,
		url,
	});
}

const ids = new Set<string>();
for (const deck of decks.values()) {
	if (ids.has(deck.id)) {
		throw new Error(`Duplicate deck id: ${deck.id}`);
	}

	ids.add(deck.id);
}

await fs.writeFile(
	manifestPath,
	JSON.stringify([...decks.values()], null, "\t") + "\n",
);

console.log(`Wrote ${String(decks.size)} decks to ${manifestPath}.`);

function compareCodePoints(a: string, b: string) {
	return a < b ? -1 : a > b ? 1 : 0;
}

function slugify(text: string) {
	return text
		.normalize("NFKD")
		.toLowerCase()
		.replaceAll(/[^a-z\d]+/g, "-")
		.replaceAll(/^-|-$/g, "");
}
