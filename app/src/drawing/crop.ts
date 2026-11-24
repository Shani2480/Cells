/**
 * Picture Format › Crop: the arithmetic behind cropping. A picture's crop says how much of the whole picture is cut
 * from each side, as fractions of it (a negative side is empty space around it, as Fit makes). The frame is the
 * object's box; the whole picture lies at `fullBox(frame, crop)`, and cropping moves one while the other stays.
 */

export type Crop = { l: number; t: number; r: number; b: number }
export type Box = { x: number; y: number; w: number; h: number }
export type CropHandle = "nw" | "n" | "ne" | "e" | "se" | "s" | "sw" | "w"

export const NO_CROP: Crop = { l: 0, t: 0, r: 0, b: 0 }

/** Where the whole picture lies for a frame showing `crop` of it. */
export function fullBox(frame: Box, crop: Crop = NO_CROP): Box {
  const w = frame.w / Math.max(0.01, 1 - crop.l - crop.r)
  const h = frame.h / Math.max(0.01, 1 - crop.t - crop.b)
  return { x: frame.x - crop.l * w, y: frame.y - crop.t * h, w, h }
}

/** The crop that shows `frame` of a picture lying at `full`. */
export function cropOf(frame: Box, full: Box): Crop {
  const clean = (v: number) => (Math.abs(v) < 1e-6 ? 0 : Math.round(v * 1e6) / 1e6)
  return {
    l: clean((frame.x - full.x) / full.w),
    t: clean((frame.y - full.y) / full.h),
    r: clean((full.x + full.w - frame.x - frame.w) / full.w),
    b: clean((full.y + full.h - frame.y - frame.h) / full.h),
  }
}

/** True when nothing is cut or added. */
export function isNoCrop(c: Crop | undefined): boolean {
  return !c || (!c.l && !c.t && !c.r && !c.b)
}

/**
 * A crop as the picture is seen when it's flipped (its left side is then on the right), and back again: the crop is
 * kept in the picture's own sides, while handles are dragged on the screen's.
 */
export function flipCrop(c: Crop, flipH = false, flipV = false): Crop {
  return {
    l: flipH ? c.r : c.l,
    r: flipH ? c.l : c.r,
    t: flipV ? c.b : c.t,
    b: flipV ? c.t : c.b,
  }
}

/**
 * The frame after a crop handle is dragged by (dx, dy): its sides move, the picture stays. Shift on a corner keeps
 * the frame's proportions. Outward past the picture is allowed (Excel adds empty space there).
 */
export function cropFrame(frame: Box, handle: CropHandle, dx: number, dy: number, keepRatio = false, min = 8): Box {
  let { x, y, w, h } = frame
  if (handle.includes("w")) {
    x += dx
    w -= dx
  }
  if (handle.includes("e")) w += dx
  if (handle.includes("n")) {
    y += dy
    h -= dy
  }
  if (handle.includes("s")) h += dy
  if (keepRatio && handle.length === 2) {
    const ratio = frame.w / frame.h
    if (w / h > ratio) w = h * ratio
    else h = w / ratio
    if (handle.includes("w")) x = frame.x + frame.w - w
    if (handle.includes("n")) y = frame.y + frame.h - h
  }
  if (w < min) {
    if (handle.includes("w")) x -= min - w
    w = min
  }
  if (h < min) {
    if (handle.includes("n")) y -= min - h
    h = min
  }
  return { x, y, w, h }
}

/** The whole picture sized from one of its corners (while cropping), keeping its proportions. */
export function scaleFull(full: Box, corner: "nw" | "ne" | "se" | "sw", dx: number, dy: number, min = 8): Box {
  // The side dragged further sets the scale.
  const kw = (full.w + (corner.includes("e") ? dx : -dx)) / full.w
  const kh = (full.h + (corner.includes("s") ? dy : -dy)) / full.h
  const k = Math.max(kw, kh, min / full.w, min / full.h)
  const w = full.w * k
  const h = full.h * k
  return {
    x: corner.includes("w") ? full.x + full.w - w : full.x,
    y: corner.includes("n") ? full.y + full.h - h : full.y,
    w,
    h,
  }
}

/** Crop › Fill: the picture covers the frame, centred, in its own proportions (what sticks out is cut). */
export function fillCrop(frame: Box, ratio: number): Crop {
  const wide = frame.w / frame.h > ratio
  const w = wide ? frame.w : frame.h * ratio
  const h = wide ? frame.w / ratio : frame.h
  return cropOf(frame, { x: frame.x + (frame.w - w) / 2, y: frame.y + (frame.h - h) / 2, w, h })
}

/** Crop › Fit: the whole picture fits in the frame, centred, in its own proportions (empty space around it). */
export function fitCrop(frame: Box, ratio: number): Crop {
  const wide = frame.w / frame.h > ratio
  const w = wide ? frame.h * ratio : frame.w
  const h = wide ? frame.h : frame.w / ratio
  return cropOf(frame, { x: frame.x + (frame.w - w) / 2, y: frame.y + (frame.h - h) / 2, w, h })
}

/**
 * Crop › Aspect Ratio: the largest part of the picture in `ratio` (width / height), from its middle. The picture
 * keeps its size and place; the frame becomes that part.
 */
export function aspectCrop(frame: Box, crop: Crop, ratio: number): { frame: Box; crop: Crop } {
  const full = fullBox(frame, crop)
  const wide = full.w / full.h > ratio
  const w = wide ? full.h * ratio : full.w
  const h = wide ? full.h : full.w / ratio
  const next = { x: full.x + (full.w - w) / 2, y: full.y + (full.h - h) / 2, w, h }
  return { frame: next, crop: cropOf(next, full) }
}

/** Excel's Aspect Ratio menu: square, portrait and landscape. */
export const ASPECT_RATIOS: { group: string; ratios: [number, number][] }[] = [
  { group: "Square", ratios: [[1, 1]] },
  {
    group: "Portrait",
    ratios: [
      [2, 3],
      [3, 4],
      [3, 5],
      [4, 5],
    ],
  },
  {
    group: "Landscape",
    ratios: [
      [3, 2],
      [4, 3],
      [5, 3],
      [5, 4],
      [16, 9],
      [16, 10],
    ],
  },
]
