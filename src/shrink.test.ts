import { decode, encode } from "jpeg-js";
import { PNG } from "pngjs";
import { describe, expect, it } from "vitest";

import { shrinkImage } from "./shrink.ts";

function png(width: number, height: number, pixels: number[]) {
	const image = new PNG({ height, width });
	image.data = Buffer.from(pixels);
	return PNG.sync.write(image);
}

function readPng(data: Buffer | undefined) {
	const {
		data: pixels,
		height,
		width,
	} = PNG.sync.read(data ?? Buffer.alloc(0));
	return { height, pixels: [...pixels], width };
}

describe(shrinkImage, () => {
	it("averages a PNG's pixels by their alpha, so transparent ones don't darken it", () => {
		const opaqueRedThenClear = [255, 0, 0, 255, 0, 0, 0, 0];
		const opaqueBlueThenClear = [0, 0, 255, 255, 0, 0, 0, 0];

		expect(
			readPng(
				shrinkImage(
					png(4, 1, [...opaqueRedThenClear, ...opaqueBlueThenClear]),
					"png",
					2,
				),
			),
		).toEqual({
			height: 1,
			pixels: [255, 0, 0, 128, 0, 0, 255, 128],
			width: 2,
		});
	});

	it("keeps a PNG that already fits at its size", () => {
		expect(
			readPng(shrinkImage(png(2, 1, [1, 2, 3, 4, 5, 6, 7, 8]), "png", 8)),
		).toEqual({ height: 1, pixels: [1, 2, 3, 4, 5, 6, 7, 8], width: 2 });
	});

	it("shrinks a JPEG to fit its longer side", () => {
		const red = encode(
			{
				data: new Uint8Array(8200 * 100 * 4).map((_, index) =>
					index % 4 === 0 || index % 4 === 3 ? 255 : 0,
				),
				height: 100,
				width: 8200,
			},
			90,
		).data;
		const { data, height, width } = decode(
			shrinkImage(red, "jpeg", 8192) ?? Buffer.alloc(0),
			{ useTArray: true },
		);
		const middle = (50 * width + 4096) * 4;

		expect({
			height,
			red: data[middle] > 240 && data[middle + 1] < 16,
			width,
		}).toEqual({ height: 99, red: true, width: 8192 });
	});

	it("returns undefined for other formats", () => {
		expect(shrinkImage(Buffer.from("GIF89a"), "gif", 8)).toBeUndefined();
	});
});
