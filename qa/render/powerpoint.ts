import { execFile } from "node:child_process";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { promisify } from "node:util";

import { ensureKerning } from "./kerning.ts";

const run = promisify(execFile);

/** PowerPoint is sandboxed: files in its own container need no file-access prompts. */
const CONTAINER = path.join(
	os.homedir(),
	"Library/Containers/com.microsoft.Powerpoint/Data/slidev2pptx-qa",
);

const LOCK = path.join(CONTAINER, "render.lock");

/** Pixels per inch to rasterize at: 144 makes a 13.33" slide 1920px wide, like Slidev's PNG export. */
export const DPI = 144;

/** How long PowerPoint may sit idle mid-export before it counts as hung. */
const HANG_AFTER = 5 * 60_000;

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

	try {
		await withLock(async () => {
			await ensureKerning(
				CONTAINER,
				(deck, output) => exportPdf(deck, output, path.basename(deck, ".pptx")),
				async () => (await powerPointProcess())?.pid,
			);
			await exportPdf(input, pdf, name);
		});
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

/** Has PowerPoint save a deck as a PDF, killing PowerPoint if it hangs along the way. */
async function exportPdf(input: string, pdf: string, name: string) {
	// AppleScript gives up on a command after 2 minutes by default, which big decks' exports take longer than.
	const exporting = run(
		"osascript",
		[
			"-e",
			`with timeout of 1800 seconds
	tell application "Microsoft PowerPoint"
		open POSIX file "${input}"
		set presentationToSave to active presentation
		save presentationToSave in POSIX file "${pdf}" as save as PDF
		close presentationToSave saving no
	end tell
end timeout`,
		],
		{ timeout: 35 * 60_000 },
	);
	const watchdog = watchForHang(() => exporting.child.kill());

	try {
		await exporting;
	} catch (error) {
		if (watchdog.killed) {
			throw new Error("PowerPoint hung, so it was killed.", { cause: error });
		}

		// Don't leave this deck open in PowerPoint for the next render to trip over.
		await run("osascript", [
			"-e",
			`with timeout of 30 seconds
	tell application "Microsoft PowerPoint"
		repeat with candidate in presentations
			if full name of candidate contains "${name}" then close candidate saving no
		end repeat
	end tell
end timeout`,
		]).catch(() => undefined);
		throw error;
	} finally {
		watchdog.stop();
	}
}

/**
 * Watches PowerPoint's CPU time while it works for us. It sometimes deadlocks
 * mid-export (idle, ignoring every AppleEvent), which would otherwise hold the
 * render lock until the export's 30 minute timeout. Once it has done nothing
 * for HANG_AFTER and only has QA files open, this kills it; the next render
 * relaunches it.
 */
function watchForHang(onKill: () => void) {
	const state = { killed: false };
	let last: undefined | { at: number; cpu: number };
	const interval = setInterval(() => {
		void (async () => {
			const powerPoint = await powerPointProcess();
			if (!powerPoint) {
				last = undefined;
				return;
			}

			if (!last || powerPoint.cpu - last.cpu > 1) {
				last = { at: Date.now(), cpu: powerPoint.cpu };
				return;
			}

			if (Date.now() - last.at < HANG_AFTER || state.killed) {
				return;
			}

			const documents = await openDocuments(powerPoint.pid);
			const theirs = documents.filter(
				(document) => !document.startsWith(CONTAINER + path.sep),
			);
			if (theirs.length) {
				console.error(
					`PowerPoint looks hung, but has other documents open, so it's left alone: ${theirs.join(", ")}`,
				);
				last = { at: Date.now(), cpu: powerPoint.cpu };
				return;
			}

			state.killed = true;
			console.error(
				`PowerPoint used no CPU for ${String(HANG_AFTER / 60_000)} minutes mid-export; killing it.`,
			);
			await run("kill", ["-9", String(powerPoint.pid)]).catch(() => undefined);
			onKill();
		})();
	}, 30_000);

	return {
		get killed() {
			return state.killed;
		},
		stop: () => {
			clearInterval(interval);
		},
	};
}

/** PowerPoint's process id and the CPU seconds it has used, if it's running. */
async function powerPointProcess() {
	const { stdout } = await run("ps", ["-axo", "pid=,time=,comm="]);
	const line = stdout
		.split("\n")
		.find((candidate) => candidate.endsWith("/MacOS/Microsoft PowerPoint"));
	const match = line?.trim().match(/^(\d+)\s+(?:(\d+):)?(\d+):([\d.]+)\s/);
	if (!match) {
		return undefined;
	}

	const [, pid, hours = "0", minutes, seconds] = match;
	return {
		cpu: Number(hours) * 3600 + Number(minutes) * 60 + Number(seconds),
		pid: Number(pid),
	};
}

/** The presentations and PDFs a process has open. */
async function openDocuments(pid: number) {
	const { stdout } = await run("lsof", ["-p", String(pid), "-Fn"]).catch(
		() => ({ stdout: "" }),
	);
	return stdout
		.split("\n")
		.filter((line) => line.startsWith("n"))
		.map((line) => line.slice(1))
		.filter((file) => /\.(?:key|pdf|pot[mx]?|pp[st][mx]?)$/i.test(file));
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
