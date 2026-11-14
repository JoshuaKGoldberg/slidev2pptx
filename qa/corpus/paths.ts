import path from "node:path";

export const corpusDirectory = import.meta.dirname;

export const decksDirectory = path.join(corpusDirectory, "decks");

export const manifestPath = path.join(corpusDirectory, "manifest.json");

export const deckPath = (id: string) => path.join(decksDirectory, `${id}.pptx`);
