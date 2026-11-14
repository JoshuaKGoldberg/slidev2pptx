import { decode, encode } from "jpeg-js";
import { PNG } from "pngjs";

interface Pixels {
	/** RGBA, 8 bits per channel. */
	data: Uint8Array;
	height: number;
	width: number;
}

/**
 * Shrinks a JPEG or PNG so its longer side is at most `side` pixels,
 * averaging the pixels each new one covers. Returns undefined for other formats.
 */
export function shrinkImage(data: Buffer, format: string, side: number) {
	switch (format) {
		case "jpeg": {
			const decoded = decode(data, {
				formatAsRGBA: true,
				maxMemoryUsageInMB: 2048,
				maxResolutionInMP: 250,
				useTArray: true,
			});
			return encode(shrink(decoded, side), 90).data;
		}

		case "png": {
			const shrunk = shrink(PNG.sync.read(data), side);
			const png = new PNG({ height: shrunk.height, width: shrunk.width });
			png.data = Buffer.from(shrunk.data);
			return PNG.sync.write(png);
		}

		default:
			return undefined;
	}
}

function shrink(source: Pixels, side: number): Pixels {
	const scale = Math.min(1, side / Math.max(source.width, source.height));
	const width = Math.max(1, Math.floor(source.width * scale));
	const height = Math.max(1, Math.floor(source.height * scale));
	const data = new Uint8Array(width * height * 4);

	// One row of new pixels at a time: per pixel, its red, green, and blue sums
	// weighted by alpha, so transparent pixels don't darken it, then its alpha
	// sum, then how many old pixels it covers.
	const sums = new Float64Array(width * 5);
	let row = 0;
	const flush = () => {
		for (let x = 0; x < width; x += 1) {
			const alpha = sums[x * 5 + 3];
			const to = (row * width + x) * 4;
			for (let channel = 0; channel < 3; channel += 1) {
				data[to + channel] = alpha
					? Math.round(sums[x * 5 + channel] / alpha)
					: 0;
			}

			data[to + 3] = Math.round(alpha / sums[x * 5 + 4]);
		}

		sums.fill(0);
	};

	for (let y = 0; y < source.height; y += 1) {
		const target = Math.min(height - 1, Math.floor(y * scale));
		if (target !== row) {
			flush();
			row = target;
		}

		for (let x = 0; x < source.width; x += 1) {
			const from = (y * source.width + x) * 4;
			const into = Math.min(width - 1, Math.floor(x * scale)) * 5;
			const alpha = source.data[from + 3];
			for (let channel = 0; channel < 3; channel += 1) {
				sums[into + channel] += source.data[from + channel] * alpha;
			}

			sums[into + 3] += alpha;
			sums[into + 4] += 1;
		}
	}

	flush();
	return { data, height, width };
}
