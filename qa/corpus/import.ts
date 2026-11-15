/**
 * Copies already-downloaded decks into qa/corpus/decks/, matching each to the
 * manifest by its text rather than trusting file names, then records each
 * deck's hash and slide count in the manifest.
 * @example
 * ```shell
 * node qa/corpus/import.ts path/to/downloaded/decks
 * ```
 */
import JSZip from "jszip";
import { createHash } from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";

import type { Deck } from "./types.ts";

import { deckPath, decksDirectory, manifestPath } from "./paths.ts";

const source = process.argv[2];
if (!source) {
	throw new Error("Usage: node qa/corpus/import.ts <directory of .pptx files>");
}

const manifest = JSON.parse(await fs.readFile(manifestPath, "utf8")) as Deck[];
await fs.mkdir(decksDirectory, { recursive: true });

const files = (await fs.readdir(source)).filter((file) =>
	file.endsWith(".pptx"),
);
const claimed = new Map<string, string>();

for (const file of files.sort()) {
	const buffer = await fs.readFile(path.join(source, file));
	const { slides, text } = await readDeck(buffer);
	const choice = manifest.find(
		(deck) => deck.number === Number.parseInt(file, 10),
	);
	if (!choice) {
		throw new Error(`${file} has no deck numbered like it.`);
	}

	// Numbers come from earlier downloads, so check each against the deck's own text.
	const score = scoreDeck(choice, text);
	const best = manifest
		.filter((deck) => deck.title !== choice.title)
		.map((deck) => ({ deck, score: scoreDeck(deck, text) }))
		.sort((a, b) => b.score - a.score)[0];
	const suspicious = score < 0.75 || best.score > score;
	console.log(
		`${suspicious ? "CHECK" : "ok   "} ${file} -> ${choice.id} (${score.toFixed(2)})${suspicious ? `; best other: ${best.deck.id} (${best.score.toFixed(2)}); text: ${text.replaceAll(/\s+/g, " ").slice(0, 160)}` : ""}`,
	);

	const previous = claimed.get(choice.id);
	if (previous) {
		throw new Error(`${file} and ${previous} both matched ${choice.id}.`);
	}

	claimed.set(choice.id, file);
	await fs.writeFile(deckPath(choice.id), buffer);
	choice.sha256 = createHash("sha256").update(buffer).digest("hex");
	choice.slides = slides;
}

await fs.writeFile(manifestPath, JSON.stringify(manifest, null, "\t") + "\n");

const missing = manifest.filter((deck) => !deck.sha256);
console.log(
	`Imported ${String(claimed.size)} decks; ${String(missing.length)} still missing:`,
);
for (const deck of missing) {
	console.log(`  #${String(deck.number)} ${deck.id} ${deck.url}`);
}

async function readDeck(buffer: Buffer) {
	const zip = await JSZip.loadAsync(buffer);
	const slideFiles = Object.keys(zip.files).filter((name) =>
		/^ppt\/slides\/slide\d+\.xml$/.test(name),
	);
	const first = slideFiles
		.sort(
			(a, b) =>
				Number(/\d+/.exec(path.basename(a))?.[0]) -
				Number(/\d+/.exec(path.basename(b))?.[0]),
		)
		.slice(0, 3);
	const texts = await Promise.all(
		[...first, "docProps/core.xml"].map(
			async (name) => (await zip.file(name)?.async("string")) ?? "",
		),
	);
	return {
		slides: slideFiles.length,
		text: texts
			.join(" ")
			.replaceAll(/<[^>]+>/g, " ")
			.replaceAll("&amp;", "&")
			.replaceAll("&apos;", "'"),
	};
}

/** How well a deck's title and, to break ties between repeat talks, its events match the text. */
function scoreDeck(deck: Deck, text: string) {
	return (
		similarity(deck.title, text) +
		Math.max(...deck.events.map(({ event }) => similarity(event, text))) / 2
	);
}

/** The fraction of the title's words that appear in the text. */
function similarity(title: string, text: string) {
	const haystack = new Set(words(text));
	const needles = words(title);
	return needles.filter((word) => haystack.has(word)).length / needles.length;
}

function words(text: string) {
	return text
		.toLowerCase()
		.replaceAll(/[’']/g, "")
		.split(/[^a-z\d]+/)
		.filter((word) => word.length > 1);
}
