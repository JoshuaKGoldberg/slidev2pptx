/**
 * Round-trips many corpus decks, a few at a time, retrying each failure once.
 * Keeps the Mac and its display awake meanwhile, as PowerPoint's PDF export
 * stalls while the display sleeps.
 * @example
 * ```shell
 * node qa/batch.ts --concurrency 3 --from convert [deck ids...]
 * ```
 */
import { execFile, spawn } from "node:child_process";
import fs from "node:fs/promises";
import path from "node:path";
import { parseArgs, promisify } from "node:util";

import type { Deck } from "./corpus/types.ts";

import { manifestPath } from "./corpus/paths.ts";

const run = promisify(execFile);

const { positionals, values } = parseArgs({
	allowPositionals: true,
	options: {
		concurrency: { default: "3", type: "string" },
		from: { default: "convert", type: "string" },
	},
});

const manifest = JSON.parse(await fs.readFile(manifestPath, "utf8")) as Deck[];
const queue = positionals.length
	? positionals
	: manifest.filter((deck) => deck.sha256).map((deck) => deck.id);
const awake = spawn("caffeinate", ["-d", "-i", "-w", String(process.pid)], {
	stdio: "ignore",
});

const failed: string[] = [];
const runs = path.join(import.meta.dirname, "runs");

async function runDeck(id: string) {
	const log = path.join(runs, `${id}.log`);
	for (let attempt = 1; attempt <= 2; attempt += 1) {
		try {
			const started = Date.now();
			const { stderr, stdout } = await run(
				process.execPath,
				[path.join(import.meta.dirname, "run.ts"), id, "--from", values.from],
				{ maxBuffer: 64 * 1024 * 1024 },
			);
			// stderr has the PowerPoint watchdog's notes, which explain slow runs.
			await fs.writeFile(log, `${stdout}\n${stderr}`);
			const minutes = Math.round((Date.now() - started) / 60_000);
			console.log(
				`${stdout.trim().split("\n").at(-1) ?? id} (${String(minutes)} min)`,
			);
			return;
		} catch (error) {
			const { stderr = "", stdout = "" } = error as {
				stderr?: string;
				stdout?: string;
			};
			// Kept apart from the deck's log, which a successful retry overwrites.
			const failure = path.join(runs, `${id}.attempt-${String(attempt)}.log`);
			await fs.writeFile(failure, `${stdout}\n${stderr}`);
			console.log(`${id}: attempt ${String(attempt)} failed; see ${failure}`);
		}
	}

	failed.push(id);
}

await Promise.all(
	Array.from({ length: Number(values.concurrency) }, async () => {
		for (let id = queue.shift(); id; id = queue.shift()) {
			await runDeck(id);
		}
	}),
);

awake.kill();
console.log(
	failed.length ? `Failed: ${failed.join(", ")}` : "Every deck round-tripped.",
);
