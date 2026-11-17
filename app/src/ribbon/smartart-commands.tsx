import {
  addSmartArtNode,
  colorCss,
  demoteSmartArtNode,
  drawnLayout,
  moveSmartArtNode,
  promoteSmartArtNode,
  SMARTART_COLORS,
  SMARTART_LAYOUTS,
  smartArtColors,
} from "@cells/core"
import { SmartArtPicture } from "@/shell/SmartArtDialog"
import type { Command, MenuEntry, RibbonEnv } from "./registry"

/**
 * Insert › SmartArt and the SmartArt Design tab: points added, moved and put up or down a level (on the Text Pane's
 * current point), another layout or colours, back to the plain drawing, or made into ordinary shapes.
 */

const has = (env: RibbonEnv) => env.c.objects.smartArt() !== null
const point = (env: RibbonEnv) => env.c.objects.smartArtPoint

const layoutsPanel: MenuEntry = {
  panel: (env, close) => {
    const sa = env.c.objects.smartArt()
    if (!sa) return null
    const current = drawnLayout(sa.smartArt.layout).id
    return (
      <div className="smartart-gallery" role="menu" aria-label="Layouts">
        {SMARTART_LAYOUTS.map((l) => (
          <button
            key={l.id}
            type="button"
            role="menuitemradio"
            aria-checked={l.id === current}
            className={`smartart-layout${l.id === current ? " is-selected" : ""}`}
            title={l.name}
            aria-label={l.name}
            onClick={() => {
              close()
              env.c.objects.setSmartArt(sa.id, { layout: l.id }, "Change Layout")
            }}
          >
            <SmartArtPicture layout={l} theme={env.c.wb.theme} w={84} h={64} colors={sa.smartArt.colors} />
          </button>
        ))}
      </div>
    )
  },
}

const colorsPanel: MenuEntry = {
  panel: (env, close) => {
    const sa = env.c.objects.smartArt()
    if (!sa) return null
    const current = smartArtColors(sa.smartArt.colors).id
    const theme = env.c.wb.theme
    return (
      <div className="smartart-colors" role="menu" aria-label="Change Colors">
        {SMARTART_COLORS.map((c) => (
          <button
            key={c.id}
            type="button"
            role="menuitemradio"
            aria-checked={c.id === current}
            className={`smartart-color${c.id === current ? " is-selected" : ""}`}
            title={c.name}
            aria-label={c.name}
            onClick={() => {
              close()
              env.c.objects.setSmartArt(sa.id, { colors: c.id }, "Change Colors")
            }}
          >
            {[0, 1, 2].map((i) => (
              <span
                key={i}
                style={{ background: colorCss({ theme: c.accents[i % c.accents.length] as never }, theme).hex }}
              />
            ))}
          </button>
        ))}
      </div>
    )
  },
}

export function smartArtCommands(): Command[] {
  return [
    {
      id: "SmartArtInsert",
      label: "SmartArt",
      description: "Insert a SmartArt graphic: a list, process, cycle or matrix drawn from text you type.",
      keytip: "M",
      run: (env) => env.dialog("smartArt"),
    },
    {
      id: "SmartArtAddShape",
      label: "Add Shape",
      description: "Add a shape to the SmartArt, after the one you're on.",
      keytip: "A",
      enabled: has,
      run: (env) => env.c.objects.editSmartArt((n, p) => addSmartArtNode(n, p), "Add Shape"),
      menu: ["SmartArtAddShapeAfter", "SmartArtAddShapeBefore"],
    },
    {
      id: "SmartArtAddShapeAfter",
      label: "Add Shape After",
      enabled: has,
      run: (env) => env.c.objects.editSmartArt((n, p) => addSmartArtNode(n, p), "Add Shape"),
    },
    {
      id: "SmartArtAddShapeBefore",
      label: "Add Shape Before",
      enabled: has,
      run: (env) => env.c.objects.editSmartArt((n, p) => addSmartArtNode(n, p, true), "Add Shape"),
    },
    {
      id: "SmartArtPromote",
      label: "Promote",
      description: "Move the point you're on up a level.",
      keytip: "P",
      enabled: (env) => has(env) && (point(env)?.length ?? 0) > 1,
      run: (env) => env.c.objects.editSmartArt((n, p) => p && promoteSmartArtNode(n, p), "Promote"),
    },
    {
      id: "SmartArtDemote",
      label: "Demote",
      description: "Move the point you're on down a level, under the one before it.",
      keytip: "D",
      enabled: (env) => has(env) && (point(env)?.at(-1) ?? 0) > 0,
      run: (env) => env.c.objects.editSmartArt((n, p) => p && demoteSmartArtNode(n, p), "Demote"),
    },
    {
      id: "SmartArtMoveUp",
      label: "Move Up",
      description: "Move the point you're on before the one above it.",
      keytip: "U",
      enabled: (env) => has(env) && (point(env)?.at(-1) ?? 0) > 0,
      run: (env) => env.c.objects.editSmartArt((n, p) => p && moveSmartArtNode(n, p, -1), "Move Up"),
    },
    {
      id: "SmartArtMoveDown",
      label: "Move Down",
      description: "Move the point you're on after the one below it.",
      keytip: "W",
      enabled: (env) => has(env) && !!point(env),
      run: (env) => env.c.objects.editSmartArt((n, p) => p && moveSmartArtNode(n, p, 1), "Move Down"),
    },
    {
      id: "SmartArtTextPane",
      label: "Text Pane",
      description: "Show or hide the pane to type the SmartArt's text in.",
      keytip: "T",
      enabled: has,
      pressed: (env) => env.c.objects.textPaneOpen,
      run: (env) => env.c.objects.setTextPane(),
    },
    {
      id: "SmartArtChangeLayoutGallery",
      label: "Change Layout",
      description: "Draw the SmartArt's text with another layout.",
      keytip: "L",
      enabled: has,
      menu: [layoutsPanel],
    },
    {
      id: "SmartArtChangeColorsGallery",
      label: "Change Colors",
      description: "Colour the SmartArt with one of the theme's colours, or each shape in turn.",
      keytip: "C",
      enabled: has,
      menu: [colorsPanel],
    },
    {
      id: "SmartArtResetGraphic",
      label: "Reset Graphic",
      description: "Discard the formatting changes made to the SmartArt; it's drawn again from its text.",
      keytip: "G",
      enabled: has,
      run: (env) => {
        const sa = env.c.objects.smartArt()
        if (sa) env.c.objects.setSmartArt(sa.id, {}, "Reset Graphic")
      },
    },
    {
      id: "SmartArtConvertToShapes",
      label: "Convert to Shapes",
      description: "Make the SmartArt a group of ordinary shapes; its text is no longer laid out for you.",
      keytip: "V",
      enabled: has,
      run: (env) => {
        const sa = env.c.objects.smartArt()
        if (sa) env.c.objects.convertToShapes(sa.id)
      },
    },
  ]
}
