import { execFile } from "node:child_process";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { promisify } from "node:util";

const run = promisify(execFile);

/** PowerPoint is sandboxed: files in its own container need no file-access prompts. */
const CONTAINER = path.join(
	os.homedir(),
	"Library/Containers/com.microsoft.Powerpoint/Data/slidev2pptx-qa",
);

const LOCK = path.join(CONTAINER, "render.lock");

/** Pixels per inch to rasterize at: 144 makes a 13.33" slide 1920px wide, like Slidev's PNG export. */
export const DPI = 144;

/**
 * Renders every slide of a .pptx to PNGs with PowerPoint itself, one deck at a
 * time, as PowerPoint is a single app shared by every caller on this machine.
 * Hidden slides aren't exported.
 */
export async function renderWithPowerPoint(pptx: string, directory: string) {
	await fs.mkdir(CONTAINER, { recursive: true });
	await fs.rm(directory, { force: true, recursive: true });
	await fs.mkdir(directory, { recursive: true });

	const name = `${String(process.pid)}-${String(Date.now())}`;
	const input = path.join(CONTAINER, `${name}.pptx`);
	const pdf = path.join(CONTAINER, `${name}.pdf`);
	await fs.copyFile(pptx, input);

	await withLock(async () => {
		await run(
			"osascript",
			[
				"-e",
				`tell application "Microsoft PowerPoint"
	open POSIX file "${input}"
	set presentationToSave to active presentation
	save presentationToSave in POSIX file "${pdf}" as save as PDF
	close presentationToSave saving no
end tell`,
			],
			{ timeout: 10 * 60_000 },
		);
	});

	try {
		await run("pdftoppm", [
			"-r",
			String(DPI),
			"-png",
			pdf,
			path.join(directory, "slide"),
		]);
	} finally {
		await fs.rm(input, { force: true });
		await fs.rm(pdf, { force: true });
	}

	return (await fs.readdir(directory))
		.filter((file) => file.endsWith(".png"))
		.sort()
		.map((file) => path.join(directory, file));
}

/** Holds a cross-process lock (an exclusive directory) while running a task. */
async function withLock(task: () => Promise<void>) {
	const started = Date.now();
	for (;;) {
		try {
			await fs.mkdir(LOCK);
			break;
		} catch {
			// A lock older than 15 minutes belongs to a crashed render.
			const stats = await fs.stat(LOCK).catch(() => undefined);
			if (stats && Date.now() - stats.mtimeMs > 15 * 60_000) {
				await fs.rm(LOCK, { force: true, recursive: true });
				continue;
			}

			if (Date.now() - started > 60 * 60_000) {
				throw new Error("Timed out waiting for PowerPoint.");
			}

			await new Promise((resolve) => setTimeout(resolve, 1000));
		}
	}

	try {
		await task();
	} finally {
		await fs.rm(LOCK, { force: true, recursive: true });
	}
}
