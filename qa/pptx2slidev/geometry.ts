/** A 2D affine transform, as in SVG's matrix(a b c d e f). */
export interface Matrix {
	a: number;
	b: number;
	c: number;
	d: number;
	e: number;
	f: number;
}

/** A shape's frame in slide px: unrotated size around its center, then rotated. */
export interface Placement {
	height: number;
	/** Clockwise degrees. */
	rotation: number;
	width: number;
	x: number;
	y: number;
}

export const IDENTITY: Matrix = { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 };

/** EMU per slide px: pptx-svg draws slides in points, which Slidev's 960px-wide canvas matches. */
export const EMU_PER_PX = 12_700;

export function multiply(first: Matrix, second: Matrix): Matrix {
	return {
		a: first.a * second.a + first.c * second.b,
		b: first.b * second.a + first.d * second.b,
		c: first.a * second.c + first.c * second.d,
		d: first.b * second.c + first.d * second.d,
		e: first.a * second.e + first.c * second.f + first.e,
		f: first.b * second.e + first.d * second.f + first.f,
	};
}

/** Parses an SVG transform attribute: translate, scale, rotate, and matrix. */
export function parseTransform(transform: null | string): Matrix {
	let result = IDENTITY;
	for (const [, name, args] of (transform ?? "").matchAll(
		/(\w+)\(([^)]*)\)/g,
	)) {
		const values = args
			.split(/[\s,]+/)
			.filter(Boolean)
			.map(Number);
		result = multiply(result, transformOf(name, values));
	}

	return result;
}

/** Where a shape with an EMU frame (in its group's child space) lands on the slide. */
export function place(
	matrix: Matrix,
	{
		cx,
		cy,
		rotation,
		x,
		y,
	}: { cx: number; cy: number; rotation: number; x: number; y: number },
): Placement {
	const width = cx / EMU_PER_PX;
	const height = cy / EMU_PER_PX;
	const centerX = x / EMU_PER_PX + width / 2;
	const centerY = y / EMU_PER_PX + height / 2;
	const scaleX = Math.hypot(matrix.a, matrix.b);
	const scaleY = Math.hypot(matrix.c, matrix.d);
	const center = {
		x: matrix.a * centerX + matrix.c * centerY + matrix.e,
		y: matrix.b * centerX + matrix.d * centerY + matrix.f,
	};
	const scaledWidth = width * scaleX;
	const scaledHeight = height * scaleY;
	return {
		height: scaledHeight,
		rotation: normalizeDegrees(
			rotation + (Math.atan2(matrix.b, matrix.a) * 180) / Math.PI,
		),
		width: scaledWidth,
		x: center.x - scaledWidth / 2,
		y: center.y - scaledHeight / 2,
	};
}

export function round(value: number) {
	return Math.round(value * 100) / 100;
}

function normalizeDegrees(degrees: number) {
	const normalized = ((degrees % 360) + 360) % 360;
	return Math.abs(normalized) < 0.01 || Math.abs(normalized - 360) < 0.01
		? 0
		: round(normalized);
}

function transformOf(name: string, values: number[]): Matrix {
	switch (name) {
		case "matrix": {
			const [a, b, c, d, e, f] = values;
			return { a, b, c, d, e, f };
		}

		case "rotate": {
			const [degrees, cx = 0, cy = 0] = values;
			const radians = (degrees * Math.PI) / 180;
			const cos = Math.cos(radians);
			const sin = Math.sin(radians);
			return multiply(
				multiply(
					{ ...IDENTITY, e: cx, f: cy },
					{ a: cos, b: sin, c: -sin, d: cos, e: 0, f: 0 },
				),
				{ ...IDENTITY, e: -cx, f: -cy },
			);
		}

		case "scale": {
			const [sx, sy = sx] = values;
			return { ...IDENTITY, a: sx, d: sy };
		}

		case "translate": {
			const [tx, ty = 0] = values;
			return { ...IDENTITY, e: tx, f: ty };
		}

		default:
			return IDENTITY;
	}
}
