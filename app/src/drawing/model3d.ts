import { model3dXml, type Model3D } from "@cells/core"

/**
 * Insert › 3D Models: a model drawn as Office draws it. Office scales the model so its longest side is a metre and
 * centres it, puts a 45° camera where the model's bounding sphere just fits, turns the model by its rotation, and
 * crops the view (a square `viewport` pixels across) to the model's outline: that outline is the object's size.
 *
 * three.js loads only when a model is drawn.
 */

/** Office's lengths: 36,000,000 to the metre (camera and offsets), angles in 60,000ths of a degree. */
export const M3D_UNIT = 36_000_000
export const M3D_ANGLE = 60_000

export type ModelSettings = {
  /** Metres per model unit, and the offset that centres the model (metres). */
  meterPerUnit: number
  preTrans: [number, number, number]
  /** Camera distance on +z (metres) and field of view (degrees). */
  cameraZ: number
  fov: number
  lights: { color: [number, number, number]; intensity: number; pos: [number, number, number] }[]
  ambient: { color: [number, number, number]; intensity: number }
}

/** Office's three lights and soft ambient light, as Excel writes them for a new model. */
export const OFFICE_LIGHTS: ModelSettings["lights"] = [
  { color: [1, 0.75, 0.5], intensity: 9.765625, pos: [0.61, 1.97, 0.454] },
  { color: [0.4, 0.6, 0.95], intensity: 12.25, pos: [-1.055, 1.42, 1.601] },
  { color: [0.86837, 0.727, 1], intensity: 3.125, pos: [-1.048, 1.613, -0.966] },
]

const num = (xml: string, re: RegExp, d: number) => {
  const m = re.exec(xml)
  return m ? Number(m[1]) : d
}
const frac = (xml: string, tag: string, d: number) => {
  const m = new RegExp(`<am3d:${tag}\\b[^>]*\\bn="(-?\\d+)"[^>]*\\bd="(\\d+)"`).exec(xml)
  return m ? Number(m[1]) / Number(m[2]) : d
}

/** The settings of a model's `am3d:model3d` element (its camera, how it's scaled and its lights). */
export function modelSettings(xml: string): ModelSettings {
  const pre = /<am3d:preTrans\b([^>]*)\/>/.exec(xml)?.[1] ?? ""
  const v = (attrs: string, name: string) => Number(new RegExp(`\\b${name}="(-?\\d+)"`).exec(attrs)?.[1] ?? 0)
  const lights = [...xml.matchAll(/<am3d:ptLight\b[\s\S]*?<\/am3d:ptLight>/g)].map(([l]) => {
    const c = /<a:scrgbClr r="(\d+)" g="(\d+)" b="(\d+)"/.exec(l)
    const p = /<am3d:pos x="(-?\d+)" y="(-?\d+)" z="(-?\d+)"/.exec(l)
    return {
      color: (c ? [c[1], c[2], c[3]].map((x) => Number(x) / 100_000) : [1, 1, 1]) as [number, number, number],
      intensity: frac(l, "intensity", 1),
      pos: (p ? [p[1], p[2], p[3]].map((x) => Number(x) / M3D_UNIT) : [0, 1, 1]) as [number, number, number],
    }
  })
  const amb = /<am3d:ambientLight>[\s\S]*?<\/am3d:ambientLight>/.exec(xml)?.[0] ?? ""
  const ac = /<a:scrgbClr r="(\d+)" g="(\d+)" b="(\d+)"/.exec(amb)
  return {
    meterPerUnit: frac(xml, "meterPerModelUnit", 1),
    preTrans: [v(pre, "dx") / M3D_UNIT, v(pre, "dy") / M3D_UNIT, v(pre, "dz") / M3D_UNIT],
    cameraZ: num(xml, /<am3d:camera>\s*<am3d:pos\b[^>]*\bz="(-?\d+)"/, 2 * M3D_UNIT) / M3D_UNIT,
    fov: num(xml, /<am3d:perspective\b[^>]*\bfov="(\d+)"/, 45 * M3D_ANGLE) / M3D_ANGLE,
    lights: lights.length ? lights : OFFICE_LIGHTS,
    ambient: {
      color: (ac ? [ac[1], ac[2], ac[3]].map((x) => Number(x) / 100_000) : [0.5, 0.5, 0.5]) as [number, number, number],
      intensity: frac(amb, "illuminance", 0.5),
    },
  }
}

type Three = typeof import("three")
let three: Promise<{
  T: Three
  GLTFLoader: typeof import("three/examples/jsm/loaders/GLTFLoader.js").GLTFLoader
}> | null = null
function loadThree() {
  three ??= Promise.all([import("three"), import("three/examples/jsm/loaders/GLTFLoader.js")]).then(([T, g]) => ({
    T,
    GLTFLoader: g.GLTFLoader,
  }))
  return three
}

const scenes = new Map<Uint8Array, Promise<import("three").Object3D>>()

/** The model's scene, read once per model data. */
async function sceneOf(data: Uint8Array) {
  let s = scenes.get(data)
  if (!s) {
    s = loadThree().then(
      ({ GLTFLoader }) =>
        new Promise((resolve, reject) => {
          const copy = data.slice().buffer
          new GLTFLoader().parse(copy, "", (gltf) => resolve(gltf.scene), reject)
        }),
    )
    scenes.set(data, s)
  }
  return s
}

/** How Office fits a model it inserts: a metre on its longest side, centred, the camera just clear of it. */
export async function fitModel(
  data: Uint8Array,
): Promise<Pick<ModelSettings, "meterPerUnit" | "preTrans" | "cameraZ">> {
  const { T } = await loadThree()
  const scene = await sceneOf(data)
  const box = new T.Box3().setFromObject(scene)
  const size = box.getSize(new T.Vector3())
  const centre = box.getCenter(new T.Vector3())
  const longest = Math.max(size.x, size.y, size.z) || 1
  const meterPerUnit = 1 / longest
  const radius = size.multiplyScalar(meterPerUnit / 2).length()
  return {
    meterPerUnit,
    preTrans: [-centre.x * meterPerUnit, -centre.y * meterPerUnit, -centre.z * meterPerUnit],
    cameraZ: radius / Math.sin(((45 / 2) * Math.PI) / 180),
  }
}

/**
 * How Office's light values turn into three.js lights: its lights shine from their direction, its ambient light is
 * strong and the picture is tone mapped. Matched to Excel's own pictures of a coloured box, within about 10 levels.
 */
const LIGHTING = { ambient: 10, directional: 0.4 }

export type ModelPicture = {
  /** The picture (PNG data URL; empty when drawn `into` a canvas), `scale` pixels per pixel of the frame. */
  src: string
  /** The model's outline in pixels: the frame that fits the model (Office's AutoFit). */
  size: { w: number; h: number }
  /** The view square it was drawn with (`fit` changes it). */
  viewport: number
}

/** One WebGL canvas for every model picture (browsers allow only a few at a time). */
let shared: { renderer: import("three").WebGLRenderer; canvas: HTMLCanvasElement } | null | undefined

function rendererOf(T: Three) {
  if (shared !== undefined) return shared
  try {
    const canvas = document.createElement("canvas")
    const renderer = new T.WebGLRenderer({ canvas, alpha: true, antialias: true, preserveDrawingBuffer: true })
    renderer.setPixelRatio(1)
    renderer.setClearColor(0x000000, 0)
    renderer.outputColorSpace = T.SRGBColorSpace
    renderer.toneMapping = T.ACESFilmicToneMapping
    canvas.addEventListener("webglcontextlost", () => (shared = undefined))
    shared = { renderer, canvas }
  } catch {
    shared = null
  }
  return shared
}

/**
 * The model as Office shows it: turned, lit and cropped to its outline, or centred in `box` (a frame of another
 * size). Drawn `into` a canvas (a live preview) or as a PNG. Null where WebGL isn't available.
 */
export async function renderModel(
  model: Model3D,
  options: { scale?: number; box?: { w: number; h: number }; fit?: boolean; into?: HTMLCanvasElement } = {},
): Promise<ModelPicture | null> {
  if (typeof document === "undefined") return null
  const { T } = await loadThree()
  const settings = modelSettings(model.xml ?? "")
  const source = await sceneOf(model.data)
  const scene = new T.Scene()
  // Pivot (turned) ← centred and scaled model.
  const pivot = new T.Group()
  const content = source.clone()
  content.scale.setScalar(settings.meterPerUnit)
  content.position.set(...settings.preTrans)
  pivot.add(content)
  const deg = Math.PI / 180
  // Office turns about the fixed x axis, then y, then z (checked against Excel's pictures).
  pivot.rotation.set(model.rot.x * deg, model.rot.y * deg, model.rot.z * deg, "ZYX")
  scene.add(pivot)
  scene.add(new T.AmbientLight(new T.Color(...settings.ambient.color), settings.ambient.intensity * LIGHTING.ambient))
  for (const l of settings.lights) {
    const light = new T.DirectionalLight(new T.Color(...l.color), l.intensity * LIGHTING.directional)
    light.position.set(...l.pos)
    scene.add(light)
  }
  const camera = new T.PerspectiveCamera(settings.fov, 1, 0.01, 1000)
  camera.position.set(0, 0, settings.cameraZ)
  camera.lookAt(0, 0, 0)
  camera.updateMatrixWorld()
  scene.updateMatrixWorld(true)

  // The outline: every vertex projected, in the view's -1…1.
  let x0 = Infinity
  let y0 = Infinity
  let x1 = -Infinity
  let y1 = -Infinity
  const v = new T.Vector3()
  pivot.traverse((o) => {
    const mesh = o as import("three").Mesh
    const pos = mesh.isMesh ? (mesh.geometry.getAttribute("position") as import("three").BufferAttribute) : null
    if (!pos) return
    for (let i = 0; i < pos.count; i++) {
      v.fromBufferAttribute(pos, i).applyMatrix4(mesh.matrixWorld).project(camera)
      x0 = Math.min(x0, v.x)
      x1 = Math.max(x1, v.x)
      y0 = Math.min(y0, v.y)
      y1 = Math.max(y1, v.y)
    }
  })
  if (!Number.isFinite(x0)) return null
  let vp = model.viewport
  let size = { w: ((x1 - x0) / 2) * vp, h: ((y1 - y0) / 2) * vp }
  // Fitted: the view square scaled so the model just fits the box.
  if (options.fit && options.box) {
    const k = Math.min(options.box.w / size.w, options.box.h / size.h)
    vp *= k
    size = { w: size.w * k, h: size.h * k }
  }
  const out = options.box ?? size
  const scale = options.scale ?? 2
  const r = rendererOf(T)
  if (!r) return null
  const w = Math.max(1, Math.round(out.w * scale))
  const h = Math.max(1, Math.round(out.h * scale))
  r.renderer.setSize(w, h, false)
  // The frame's part of the view square: the outline, centred in the frame.
  const left = ((x0 + 1) / 2) * vp - (out.w - size.w) / 2
  const top = ((1 - y1) / 2) * vp - (out.h - size.h) / 2
  camera.setViewOffset(vp, vp, left, top, out.w, out.h)
  r.renderer.render(scene, camera)
  if (options.into) {
    const target = options.into
    if (target.width !== w || target.height !== h) {
      target.width = w
      target.height = h
    }
    const g = target.getContext("2d")
    g?.clearRect(0, 0, w, h)
    g?.drawImage(r.canvas, 0, 0)
    return { src: "", size, viewport: vp }
  }
  return { src: r.canvas.toDataURL("image/png"), size, viewport: vp }
}

/** The model files Insert › 3D Models takes (as Excel: glTF, FBX, OBJ, 3MF, PLY and STL). */
export const MODEL_EXTENSIONS = ["glb", "gltf", "fbx", "obj", "3mf", "ply", "stl"]

/**
 * A model file as a .glb (Office keeps every model as glTF binary): a .glb as it is, other formats read with
 * three.js and written out as glTF. A .gltf must hold its data (no separate files).
 */
export async function toGlb(bytes: Uint8Array, name: string): Promise<Uint8Array> {
  const ext = name.split(".").pop()!.toLowerCase()
  if (ext === "glb") return bytes
  const { T } = await loadThree()
  const buffer = bytes.slice().buffer
  const text = () => new TextDecoder().decode(bytes)
  let object: import("three").Object3D
  if (ext === "gltf") {
    const { GLTFLoader } = await loadThree()
    object = await new Promise((resolve, reject) => new GLTFLoader().parse(text(), "", (g) => resolve(g.scene), reject))
  } else if (ext === "obj") {
    const { OBJLoader } = await import("three/examples/jsm/loaders/OBJLoader.js")
    object = new OBJLoader().parse(text())
  } else if (ext === "fbx") {
    const { FBXLoader } = await import("three/examples/jsm/loaders/FBXLoader.js")
    object = new FBXLoader().parse(buffer, "")
  } else if (ext === "3mf") {
    const { ThreeMFLoader } = await import("three/examples/jsm/loaders/3MFLoader.js")
    object = new ThreeMFLoader().parse(buffer)
  } else if (ext === "stl" || ext === "ply") {
    const geometry =
      ext === "stl"
        ? new (await import("three/examples/jsm/loaders/STLLoader.js")).STLLoader().parse(buffer)
        : new (await import("three/examples/jsm/loaders/PLYLoader.js")).PLYLoader().parse(buffer)
    if (!geometry.getAttribute("normal")) geometry.computeVertexNormals()
    const colored = !!geometry.getAttribute("color")
    // A plain model gets Office's light grey.
    const material = new T.MeshStandardMaterial({
      color: colored ? 0xffffff : 0xb4b4b4,
      vertexColors: colored,
      roughness: 0.6,
    })
    object = new T.Mesh(geometry, material)
  } else throw new Error(`Cells can't read .${ext} models. Use a .glb, .gltf, .fbx, .obj, .3mf, .ply or .stl file.`)
  const { GLTFExporter } = await import("three/examples/jsm/exporters/GLTFExporter.js")
  const out = await new GLTFExporter().parseAsync(object, { binary: true })
  return new Uint8Array(out as ArrayBuffer)
}

/** A 3D model file picked on this device, as a .glb. */
export function pickModel(): Promise<{ data: Uint8Array; name: string } | null> {
  return new Promise((resolve, reject) => {
    const input = document.createElement("input")
    input.type = "file"
    input.accept = MODEL_EXTENSIONS.map((e) => `.${e}`).join(",")
    input.onchange = () => {
      const file = input.files?.[0]
      if (!file) return resolve(null)
      void file
        .arrayBuffer()
        .then((b) => toGlb(new Uint8Array(b), file.name))
        .then((data) => resolve({ data, name: file.name }), reject)
    }
    input.oncancel = () => resolve(null)
    input.click()
  })
}

/** Office's view of a model it inserts: the `am3d` settings for `data`, with the view square it starts with. */
export async function newModel(data: Uint8Array): Promise<Model3D> {
  const fit = await fitModel(data)
  return { data, rot: { x: 0, y: 0, z: 0 }, viewport: NEW_VIEWPORT, xml: model3dXml(fit) }
}

/** The view square of a model Excel inserts: 3 inches. */
export const NEW_VIEWPORT = 288
