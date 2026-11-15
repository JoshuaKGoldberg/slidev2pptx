import JSZip from "jszip";
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

/** How long an export may take by default: original talks full of media take many minutes. */
const EXPORT_LIMIT = 30 * 60_000;

/**
 * Renders every slide of a .pptx to PNGs with PowerPoint itself, one deck at a
 * time, as PowerPoint is a single app shared by every caller on this machine.
 * Hidden slides aren't exported. PowerPoint is killed and the export tried
 * once more if it hangs or takes longer than `limit` milliseconds.
 */
export async function renderWithPowerPoint(
	pptx: string,
	directory: string,
	{ limit = EXPORT_LIMIT } = {},
) {
	await fs.mkdir(CONTAINER, { recursive: true });
	await fs.rm(directory, { force: true, recursive: true });
	await fs.mkdir(directory, { recursive: true });

	const name = `${String(process.pid)}-${String(Date.now())}`;
	const input = path.join(CONTAINER, `${name}.pptx`);
	const pdf = path.join(CONTAINER, `${name}.pdf`);
	await fs.copyFile(pptx, input);

	try {
		const attempt = async () => {
			await ensureKerning(
				CONTAINER,
				(deck, output) =>
					exportPdf(deck, output, path.basename(deck, ".pptx"), EXPORT_LIMIT),
				async () => (await powerPointProcess())?.pid,
			);
			await exportPdf(input, pdf, name, limit);
		};

		await withLock(() =>
			// A relaunched PowerPoint gets its kerning checked again.
			attempt().catch(async (error: unknown) => {
				if (!(error instanceof PowerPointKilled)) {
					throw error;
				}

				console.error(`${error.message} Trying ${pptx} once more.`);
				await attempt();
			}),
		);
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

/**
 * Renders some of a .pptx's shown slides to PNGs, by rendering a copy of it
 * with every other slide hidden.
 * @param pptx The presentation to render.
 * @param indices Positions among the shown slides, in the order to render them.
 * @param directory Where to write the PNGs.
 * @returns One PNG per index, in order.
 */
export async function renderSlidesWithPowerPoint(
	pptx: string,
	indices: number[],
	directory: string,
) {
	const zip = await JSZip.loadAsync(await fs.readFile(pptx));
	const slides = await shownSlides(zip);
	const kept = new Set(indices.map((index) => slides[index]));
	for (const slide of slides) {
		if (!kept.has(slide)) {
			const xml = await zip.file(slide)?.async("string");
			if (xml) {
				zip.file(slide, xml.replace(/<p:sld\b/, '<p:sld show="0"'));
			}
		}
	}

	const copy = `${directory}.pptx`;
	await fs.mkdir(path.dirname(copy), { recursive: true });
	await fs.writeFile(copy, await zip.generateAsync({ type: "nodebuffer" }));
	try {
		// Only a few slides are shown, so this takes PowerPoint a minute or two at most.
		const rendered = await renderWithPowerPoint(copy, directory, {
			limit: 5 * 60_000,
		});
		const order = [...kept].sort(
			(a, b) => slides.indexOf(a) - slides.indexOf(b),
		);
		return indices.map((index) => rendered[order.indexOf(slides[index])]);
	} finally {
		await fs.rm(copy, { force: true });
	}
}

/** The part names of a presentation's shown slides, in order. */
async function shownSlides(zip: JSZip) {
	const presentation = await zip.file("ppt/presentation.xml")?.async("string");
	const relationships = await zip
		.file("ppt/_rels/presentation.xml.rels")
		?.async("string");
	const targets = new Map(
		[
			...(relationships ?? "").matchAll(
				/<Relationship\b[^>]*?\sId="([^"]+)"[^>]*?\sTarget="([^"]+)"/g,
			),
		].map(([, id, target]) => [id, path.posix.join("ppt", target)]),
	);
	const slides = [];
	for (const [, id] of (presentation ?? "").matchAll(
		/<p:sldId\b[^>]*?\sr:id="([^"]+)"/g,
	)) {
		const slide = targets.get(id);
		const xml = slide && (await zip.file(slide)?.async("string"));
		if (slide && xml && !/<p:sld\b[^>]*\sshow="(?:0|false)"/.test(xml)) {
			slides.push(slide);
		}
	}

	return slides;
}

/** PowerPoint hung or stalled mid-export, and was killed. */
class PowerPointKilled extends Error {}

/** Has PowerPoint save a deck as a PDF, killing PowerPoint if it hangs or stalls along the way. */
async function exportPdf(
	input: string,
	pdf: string,
	name: string,
	limit: number,
) {
	// AppleScript gives up on a command after 2 minutes by default, which big decks' exports take longer than.
	const exporting = run(
		"osascript",
		[
			"-e",
			`with timeout of ${String(Math.ceil(limit / 1000) + 300)} seconds
	tell application "Microsoft PowerPoint"
		open POSIX file "${input}"
		set presentationToSave to active presentation
		save presentationToSave in POSIX file "${pdf}" as save as PDF
		close presentationToSave saving no
	end tell
end timeout`,
		],
		{ timeout: limit + 5 * 60_000 },
	);
	const watchdog = watchForHang(limit, () => exporting.child.kill());

	try {
		await exporting;
	} catch (error) {
		if (watchdog.killed) {
			throw new PowerPointKilled(`PowerPoint ${watchdog.killed}.`, {
				cause: error,
			});
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
 * Watches PowerPoint while it works for us. It sometimes deadlocks mid-export
 * (idle, ignoring every AppleEvent), and sometimes stalls (busy, as if behind
 * an alert no one sees, also ignoring AppleEvents), either of which would
 * otherwise hold the render lock until the export times out. Once it has done
 * nothing for HANG_AFTER, or exported for longer than `limit`, and only has QA
 * files open, this kills it; the next render relaunches it.
 */
function watchForHang(limit: number, onKill: () => void) {
	const started = Date.now();
	const state: { killed?: "hung" | "stalled" } = {};
	let last: undefined | { at: number; cpu: number };
	const interval = setInterval(() => {
		void (async () => {
			const powerPoint = await powerPointProcess();
			if (!powerPoint) {
				last = undefined;
				return;
			}

			const stalled = Date.now() - started > limit;
			if (!stalled && (!last || powerPoint.cpu - last.cpu > 1)) {
				last = { at: Date.now(), cpu: powerPoint.cpu };
				return;
			}

			if (
				state.killed ||
				(!stalled && last && Date.now() - last.at < HANG_AFTER)
			) {
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

			state.killed = stalled ? "stalled" : "hung";
			console.error(
				stalled
					? `PowerPoint took over ${String(limit / 60_000)} minutes to export; killing it.`
					: `PowerPoint used no CPU for ${String(HANG_AFTER / 60_000)} minutes mid-export; killing it.`,
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
	const owner = path.join(LOCK, "owner");
	const started = Date.now();
	for (;;) {
		try {
			await fs.mkdir(LOCK);
			await fs.writeFile(owner, String(process.pid));
			break;
		} catch {
			if (await abandoned(owner)) {
				await fs.rm(LOCK, { force: true, recursive: true });
				continue;
			}

			if (Date.now() - started > 2 * 60 * 60_000) {
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

/** Whether the lock's owner has exited without releasing it. */
async function abandoned(owner: string) {
	const pid = Number(await fs.readFile(owner, "utf8").catch(() => ""));
	if (!pid) {
		// The owner writes its pid right after taking the lock; give it a moment.
		const stats = await fs.stat(LOCK).catch(() => undefined);
		return !!stats && Date.now() - stats.mtimeMs > 30_000;
	}

	try {
		process.kill(pid, 0);
		return false;
	} catch (error) {
		return (error as NodeJS.ErrnoException).code === "ESRCH";
	}
}
