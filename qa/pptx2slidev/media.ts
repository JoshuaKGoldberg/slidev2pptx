import { imageSize } from "image-size";

/** An image's size in pixels and its resolution, which PowerPoint sizes tiled picture fills by. */
export interface ImageMetrics {
	dpiX: number;
	dpiY: number;
	height: number;
	width: number;
}

/** PowerPoint's resolution for images that don't record one. */
const DEFAULT_DPI = 96;

/** An image's pixel size and resolution, from JPEG JFIF or EXIF headers or a PNG pHYs chunk; undefined if it can't be read. */
export function imageMetrics(buffer: Buffer): ImageMetrics | undefined {
	let size;
	try {
		size = imageSize(buffer);
	} catch {
		return undefined;
	}

	if (!size.width || !size.height) {
		return undefined;
	}

	const [dpiX, dpiY] =
		size.type === "jpg"
			? jpegResolution(buffer)
			: size.type === "png"
				? pngResolution(buffer)
				: [];
	return {
		dpiX: dpiX || DEFAULT_DPI,
		dpiY: dpiY || dpiX || DEFAULT_DPI,
		height: size.height,
		width: size.width,
	};
}

function jpegResolution(buffer: Buffer): number[] {
	let exif: number[] = [];
	for (let offset = 2; offset + 4 <= buffer.length;) {
		if (buffer[offset] !== 0xff) {
			break;
		}

		const marker = buffer[offset + 1];
		const length = buffer.readUInt16BE(offset + 2);
		const segment = buffer.subarray(offset + 4, offset + 2 + length);
		if (marker === 0xe0 && segment.toString("latin1", 0, 5) === "JFIF\0") {
			const units = segment[7];
			const x = segment.readUInt16BE(8);
			const y = segment.readUInt16BE(10);
			if (units === 1) {
				return [x, y];
			}

			if (units === 2) {
				return [x * 2.54, y * 2.54];
			}
		} else if (
			marker === 0xe1 &&
			segment.toString("latin1", 0, 6) === "Exif\0\0"
		) {
			exif = exifResolution(segment.subarray(6));
		} else if (marker === 0xda) {
			break;
		}

		offset += 2 + length;
	}

	return exif;
}

/** XResolution and YResolution from a TIFF header's first directory, in dots per inch. */
function exifResolution(tiff: Buffer): number[] {
	if (tiff.length < 8) {
		return [];
	}

	const little = tiff.toString("latin1", 0, 2) === "II";
	const u16 = (at: number) =>
		little ? tiff.readUInt16LE(at) : tiff.readUInt16BE(at);
	const u32 = (at: number) =>
		little ? tiff.readUInt32LE(at) : tiff.readUInt32BE(at);
	const directory = u32(4);
	if (directory + 2 > tiff.length) {
		return [];
	}

	const values = new Map<number, number>();
	const count = u16(directory);
	for (let index = 0; index < count; index += 1) {
		const entry = directory + 2 + index * 12;
		if (entry + 12 > tiff.length) {
			break;
		}

		const tag = u16(entry);
		const type = u16(entry + 2);
		if (type === 5) {
			const at = u32(entry + 8);
			if (at + 8 <= tiff.length && u32(at + 4)) {
				values.set(tag, u32(at) / u32(at + 4));
			}
		} else if (type === 3) {
			values.set(tag, u16(entry + 8));
		}
	}

	const scale = values.get(0x0128) === 3 ? 2.54 : 1;
	const x = values.get(0x011a);
	const y = values.get(0x011b);
	return x ? [x * scale, (y ?? x) * scale] : [];
}

function pngResolution(buffer: Buffer): number[] {
	for (let offset = 8; offset + 12 <= buffer.length;) {
		const length = buffer.readUInt32BE(offset);
		const type = buffer.toString("latin1", offset + 4, offset + 8);
		if (type === "pHYs" && buffer[offset + 16] === 1) {
			return [
				buffer.readUInt32BE(offset + 8) * 0.0254,
				buffer.readUInt32BE(offset + 12) * 0.0254,
			];
		}

		if (type === "IDAT") {
			break;
		}

		offset += 12 + length;
	}

	return [];
}
