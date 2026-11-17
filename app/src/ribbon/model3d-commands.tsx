import { useEffect, useState } from "react"
import type { Model3D } from "@cells/core"
import { model3dXml } from "@cells/core"
import { fitModel, NEW_VIEWPORT, newModel, pickModel, renderModel } from "@/drawing/model3d"
import type { Command, MenuEntry, RibbonEnv } from "./registry"

/**
 * Insert › 3D Models and the 3D Model tab: a model from a file (glTF, FBX, OBJ, 3MF, PLY, STL), turned by the
 * handle in its middle, by Office's views or by degrees, and reset.
 */

/** Office's 3D Model Views: the model seen from each side and from above or below a corner (degrees x, y, z). */
export const MODEL_VIEWS: { name: string; rot: Model3D["rot"] }[] = [
  { name: "Front", rot: { x: 0, y: 0, z: 0 } },
  { name: "Back", rot: { x: 0, y: 180, z: 0 } },
  { name: "Left", rot: { x: 0, y: 90, z: 0 } },
  { name: "Right", rot: { x: 0, y: 270, z: 0 } },
  { name: "Top", rot: { x: 90, y: 0, z: 0 } },
  { name: "Bottom", rot: { x: 270, y: 0, z: 0 } },
  { name: "Above Front Left", rot: { x: 25, y: 35, z: 0 } },
  { name: "Above Front Right", rot: { x: 25, y: 325, z: 0 } },
  { name: "Above Back Left", rot: { x: 25, y: 145, z: 0 } },
  { name: "Above Back Right", rot: { x: 25, y: 215, z: 0 } },
  { name: "Below Front Left", rot: { x: 335, y: 35, z: 0 } },
  { name: "Below Front Right", rot: { x: 335, y: 325, z: 0 } },
]

const hasModel = (env: RibbonEnv) => env.c.objects.model() !== null

/** Picks a model file and puts it in (`replace`: in place of the selected model's, keeping its view). */
async function fromFile(env: RibbonEnv, replace = false) {
  let picked: Awaited<ReturnType<typeof pickModel>>
  try {
    picked = await pickModel()
  } catch (e) {
    env.c.shell.alert(e instanceof Error ? e.message : "Cells couldn't read that 3D model.")
    return
  }
  if (!picked) return
  const failed = () => env.c.shell.alert("Cells can't show 3D models here: this device's browser has no WebGL.")
  const current = env.c.objects.model()
  if (replace && current) {
    const xml = model3dXml(await fitModel(picked.data))
    if (!(await env.c.objects.setModelView(current.id, { data: picked.data, xml }, "Change 3D Model"))) failed()
    return
  }
  if (!(await env.c.objects.insertModel3D(await newModel(picked.data)))) failed()
}

/** A view of the selected model, drawn small for the gallery. */
function ViewPicture({ model, rot, label }: { model: Model3D; rot: Model3D["rot"]; label: string }) {
  const [src, setSrc] = useState<string | null>(null)
  useEffect(() => {
    let live = true
    // Drawn one after another: the gallery's pictures share one renderer.
    void renderModel({ ...model, rot }, { box: { w: 64, h: 52 }, fit: true, scale: 2 }).then((p) => {
      if (live && p) setSrc(p.src)
    })
    return () => {
      live = false
    }
  }, [model, rot])
  return src ? (
    <img src={src} alt="" width={64} height={52} />
  ) : (
    <span className="model-view-blank" aria-label={label} />
  )
}

const viewsPanel: MenuEntry = {
  panel: (env, close) => {
    const m = env.c.objects.model()
    if (!m) return null
    return (
      <div className="model-views" role="menu" aria-label="3D Model Views">
        {MODEL_VIEWS.map((v) => (
          <button
            key={v.name}
            type="button"
            role="menuitem"
            className="model-view"
            title={v.name}
            aria-label={v.name}
            onClick={() => {
              close()
              void env.c.objects.setModelView(m.id, { rot: v.rot }, "3D Model Views")
            }}
          >
            <ViewPicture model={m.model3d} rot={v.rot} label={v.name} />
          </button>
        ))}
      </div>
    )
  },
}

/** A field for the model's turn about one axis, in degrees (Format 3D Model › Rotation). */
function rotationField(axis: "x" | "y" | "z"): Command {
  const name = axis.toUpperCase()
  return {
    id: `Model3DRotation${name}`,
    label: `${name} Rotation`,
    description: `Turn the 3D model about its ${name} axis, in degrees.`,
    enabled: hasModel,
    field: {
      width: 52,
      value: (env) => {
        const m = env.c.objects.model()
        return m ? `${Math.round(m.model3d.rot[axis] * 10) / 10}°` : ""
      },
      commit: (env, text) => {
        const m = env.c.objects.model()
        const n = parseFloat(text.replace(",", "."))
        if (!m || !Number.isFinite(n)) return
        void env.c.objects.setModelView(m.id, { rot: { ...m.model3d.rot, [axis]: n } }, "3D Rotation")
      },
    },
  }
}

export function model3dCommands(): Command[] {
  return [
    {
      id: "Insert3DModelDropdown",
      label: "3D Models",
      description:
        "Insert a 3D model you can turn to show it from any side. Cells reads .glb, .gltf, .fbx, .obj, .3mf, .ply and .stl files.",
      keytip: "3",
      run: (env) => void fromFile(env),
      menu: ["Model3DInsertFromFile"],
    },
    {
      id: "Model3DInsertFromFile",
      label: "This Device...",
      description: "Insert a 3D model from a file on this device.",
      run: (env) => void fromFile(env),
    },
    {
      id: "Model3DChange",
      label: "Change 3D Model",
      description: "Put another model from a file in place of this one, keeping its place and view.",
      keytip: "C",
      enabled: hasModel,
      run: (env) => void fromFile(env, true),
      menu: ["Model3DChangeFromFile"],
    },
    {
      id: "Model3DChangeFromFile",
      label: "This Device...",
      description: "Replace the model with one from a file on this device.",
      enabled: hasModel,
      run: (env) => void fromFile(env, true),
    },
    {
      id: "Model3DReset",
      label: "Reset 3D Model",
      description: "Discard the turns made to this 3D model.",
      keytip: "R",
      enabled: hasModel,
      run: (env) => {
        const m = env.c.objects.model()
        if (m) void env.c.objects.setModelView(m.id, { rot: { x: 0, y: 0, z: 0 } }, "Reset 3D Model", true)
      },
      menu: ["Model3DResetView", "Model3DResetViewAndSize"],
    },
    {
      id: "Model3DResetView",
      label: "Reset 3D Model",
      description: "Discard the turns made to this 3D model; it keeps its size.",
      enabled: hasModel,
      run: (env) => {
        const m = env.c.objects.model()
        if (m) void env.c.objects.setModelView(m.id, { rot: { x: 0, y: 0, z: 0 } }, "Reset 3D Model", true)
      },
    },
    {
      id: "Model3DResetViewAndSize",
      label: "Reset 3D Model and Size",
      description: "Discard the turns and the size changes made to this 3D model.",
      enabled: hasModel,
      run: (env) => {
        const m = env.c.objects.model()
        if (m)
          void env.c.objects.setModelView(
            m.id,
            { rot: { x: 0, y: 0, z: 0 }, viewport: NEW_VIEWPORT },
            "Reset 3D Model and Size",
          )
      },
    },
    {
      id: "Model3DViewsGallery",
      label: "3D Model Views",
      description: "See the model from the front, a side, above or below.",
      keytip: "V",
      enabled: hasModel,
      menu: [viewsPanel],
    },
    rotationField("x"),
    rotationField("y"),
    rotationField("z"),
  ]
}
