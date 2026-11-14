import { execFile } from "node:child_process";
import fs from "node:fs/promises";
import path from "node:path";
import { promisify } from "node:util";

const run = promisify(execFile);

const SLIDEV = path.join(
	import.meta.dirname,
	"../../node_modules/@slidev/cli/bin/slidev.mjs",
);

/**
 * Renders a Slidev deck to PNGs in Chromium, one per click state, named like
 * 001-01.png for slide 1 before any clicks. Hidden slides aren't exported.
 */
export async function renderWithSlidev(entry: string, directory: string) {
	// Concurrent exports can catch Vite re-optimizing dependencies and export nothing, so retry those.
	for (let attempt = 1; ; attempt += 1) {
		const result = await exportOnce(entry, directory);
		if (result.pngs.length || attempt === 3) {
			if (!result.pngs.length) {
				throw new Error(
					`Slidev exported no slides from ${entry}: ${result.errors.join("\n")}`,
				);
			}

			return result;
		}
	}
}

async function exportOnce(entry: string, directory: string) {
	await fs.rm(directory, { force: true, recursive: true });
	const { stderr, stdout } = await run(
		process.execPath,
		[
			SLIDEV,
			"export",
			path.basename(entry),
			"--format",
			"png",
			"--with-clicks",
			"--output",
			directory,
			"--timeout",
			"120000",
		],
		{
			cwd: path.dirname(entry),
			maxBuffer: 64 * 1024 * 1024,
			timeout: 30 * 60_000,
		},
	);
	const errors = `${stdout}\n${stderr}`
		.split("\n")
		.filter((line) => /error/i.test(line) && !/^\s+at /.test(line));

	return {
		errors: [...new Set(errors)],
		pngs: (await fs.readdir(directory).catch(() => []))
			.filter((file) => file.endsWith(".png"))
			.sort()
			.map((file) => path.join(directory, file)),
	};
}
