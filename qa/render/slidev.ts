import { parse } from "@slidev/parser";
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
	// Under load, exports can catch Vite re-optimizing dependencies (and export nothing),
	// time out waiting for a page to show, or stop partway, so retry those.
	const expected = await shownSlides(entry);
	const failures: string[] = [];
	for (let attempt = 1; attempt <= 3; attempt += 1) {
		try {
			const result = await exportOnce(entry, directory);
			const exported = new Set(
				result.pngs.map((file) => Number(path.basename(file).slice(0, 3))),
			);
			if (exported.size >= expected) {
				return result;
			}

			failures.push(
				`Exported ${String(exported.size)} of ${String(expected)} slides: ${result.errors.join("\n")}`,
			);
		} catch (error) {
			failures.push(String(error).split("\n").slice(0, 3).join("\n"));
		}
	}

	throw new Error(
		`Slidev couldn't export ${entry}:\n${failures.join("\n---\n")}`,
	);
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

/** How many slides a deck shows: those that aren't hidden or disabled. */
async function shownSlides(entry: string) {
	const { slides } = await parse(await fs.readFile(entry, "utf8"), entry);
	return slides.filter(
		({ frontmatter }) => !frontmatter.hide && !frontmatter.disabled,
	).length;
}
