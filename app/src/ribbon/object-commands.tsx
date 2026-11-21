import { colorCss, isChart, type DashStyle, type ShapeColor, type SheetShape } from "@cells/core"
import { presetGeometry, SHAPE_GALLERY } from "@/drawing/geometry"
import { pickPicture } from "@/drawing/pictures"
import { WORDART_STYLES } from "@/drawing/wordart"
import { themeEffectFilter } from "@/drawing/render"
import { applyShapeStyle, SHAPE_STYLES, type ShapeStyle } from "@/drawing/shape-styles"
import { ColorPanel, shapeColorOf } from "./colors"
import type { Command, MenuEntry, RibbonEnv } from "./registry"

/** Excel's names for the shapes in the gallery's tooltips. */
const SHAPE_LABELS: Record<string, string> = {
  line: "Line",
  straightConnector1: "Line Arrow",
  rect: "Rectangle",
  roundRect: "Rectangle: Rounded Corners",
  snip1Rect: "Rectangle: Single Corner Snipped",
  ellipse: "Oval",
  triangle: "Isosceles Triangle",
  rtTriangle: "Right Triangle",
  parallelogram: "Parallelogram",
  trapezoid: "Trapezoid",
  diamond: "Diamond",
  pentagon: "Regular Pentagon",
  hexagon: "Hexagon",
  octagon: "Octagon",
  plus: "Cross",
  can: "Cylinder",
  donut: "Circle: Hollow",
  frame: "Frame",
  heart: "Heart",
  smileyFace: "Smiley Face",
  lightningBolt: "Lightning Bolt",
  sun: "Sun",
  moon: "Moon",
  cloud: "Cloud",
  rightArrow: "Arrow: Right",
  leftArrow: "Arrow: Left",
  upArrow: "Arrow: Up",
  downArrow: "Arrow: Down",
  leftRightArrow: "Arrow: Left-Right",
  chevron: "Arrow: Chevron",
  homePlate: "Arrow: Pentagon",
  flowChartProcess: "Flowchart: Process",
  flowChartDecision: "Flowchart: Decision",
  flowChartTerminator: "Flowchart: Terminator",
  flowChartAlternateProcess: "Flowchart: Alternate Process",
  star4: "Star: 4 Points",
  star5: "Star: 5 Points",
  star6: "Star: 6 Points",
  wedgeRectCallout: "Speech Bubble: Rectangle",
  wedgeRoundRectCallout: "Speech Bubble: Rectangle with Corners Rounded",
  wedgeEllipseCallout: "Speech Bubble: Oval",
  cloudCallout: "Thought Bubble: Cloud",
}

/** A shape's outline drawn small, for the gallery. */
export function ShapeIcon({ geometry, size = 20 }: { geometry: string; size?: number }) {
  const g = presetGeometry(geometry, size - 4, size - 4, geometry.includes("Callout") ? { adj2: 40000 } : undefined)
  const end = geometry === "straightConnector1"
  return (
    <svg width={size} height={size} viewBox={`-2 -2 ${size} ${size}`} aria-hidden className="shape-icon">
      {end ? (
        <defs>
          <marker
            id="gallery-arrow"
            viewBox="0 0 10 10"
            refX="9"
            refY="5"
            markerWidth="5"
            markerHeight="5"
            orient="auto"
          >
            <path d="M0 0 L10 5 L0 10 Z" fill="currentColor" />
          </marker>
        </defs>
      ) : null}
      {[...g.paths, ...(g.lines ?? [])].map((d, i) => (
        <path
          key={i}
          d={d}
          fill="none"
          fillRule="evenodd"
          stroke="currentColor"
          strokeWidth="1.1"
          markerEnd={end ? "url(#gallery-arrow)" : undefined}
        />
      ))}
    </svg>
  )
}

const shapesPanel: MenuEntry = {
  panel: (env, close) => (
    <div className="shape-gallery" role="menu" aria-label="Shapes">
      {SHAPE_GALLERY.map((g) => (
        <div key={g.group}>
          <div className="color-heading">{g.group}</div>
          <div className="shape-gallery-row">
            {g.shapes.map((s) => (
              <button
                key={s}
                type="button"
                role="menuitem"
                className="shape-gallery-item"
                title={SHAPE_LABELS[s] ?? s}
                aria-label={SHAPE_LABELS[s] ?? s}
                onClick={() => {
                  close()
                  env.c.objects.startShape(s)
                }}
              >
                <ShapeIcon geometry={s} />
              </button>
            ))}
          </div>
        </div>
      ))}
    </div>
  ),
}

const wordArtPanel: MenuEntry = {
  panel: (env, close) => (
    <div className="wordart-gallery" role="menu" aria-label="WordArt">
      {WORDART_STYLES.map((w) => {
        const fill = w.fill?.kind === "solid" ? colorCss(w.fill.color, env.c.wb.theme).hex : "transparent"
        const outline = w.effects?.outline?.color ? colorCss(w.effects.outline.color, env.c.wb.theme).hex : undefined
        return (
          <button
            key={w.name}
            type="button"
            role="menuitem"
            className="wordart-gallery-item"
            title={w.name}
            aria-label={w.name}
            onClick={() => {
              close()
              env.c.objects.insertWordArt({ wordArt: { fill: w.fill, effects: w.effects } })
              env.c.shell.showTab?.("TabDrawingToolsFormat")
            }}
          >
            <span
              style={{
                color: fill,
                WebkitTextStroke: outline ? `1px ${outline}` : undefined,
                textShadow: [
                  w.effects?.shadow ? "2px 2px 2px rgb(0 0 0 / 0.4)" : "",
                  w.effects?.glow ? `0 0 6px ${colorCss(w.effects.glow, env.c.wb.theme).hex}` : "",
                ]
                  .filter(Boolean)
                  .join(", "),
              }}
            >
              A
            </span>
          </button>
        )
      })}
    </div>
  ),
}

const hasDrawing = (env: RibbonEnv) => env.c.objects.drawings().length > 0
const hasObjects = (env: RibbonEnv) => env.c.objects.selected.length > 0
const hasShape = (env: RibbonEnv) => env.c.objects.drawings().some((o) => o.kind === "shape")
const shapes = (env: RibbonEnv) => env.c.objects.drawings().filter((o): o is SheetShape => o.kind === "shape")

/** A style's picture in the gallery: a rounded box in its fill, outline and effect, with "Abc" in its text colour. */
function styleTile(env: RibbonEnv, s: ShapeStyle, id: string): string {
  const theme = env.c.wb.theme
  const fill = s.fill.kind === "solid" ? colorCss(s.fill.color, theme).hex : "none"
  const line = s.line.color ? colorCss(s.line.color, theme).hex : "none"
  const fx = s.effect ? themeEffectFilter(id, s.effect, theme, fill) : ""
  return (
    `${fx ? `<defs>${fx}</defs>` : ""}<rect x="4" y="4" width="48" height="28" rx="3" fill="${fill}" stroke="${line}"` +
    ` stroke-width="${Math.max(0.75, s.line.width)}"${fx ? ` filter="url(#${id})"` : ""}/>` +
    `<text x="28" y="22" text-anchor="middle" font-size="11" fill="${colorCss(s.text, theme).hex}">Abc</text>`
  )
}

const shapeStylesPanel: MenuEntry = {
  panel: (env, close) => (
    <div className="shape-styles" role="menu" aria-label="Shape Styles">
      <span className="menu-heading">Theme Styles</span>
      <div className="shape-styles-grid">
        {SHAPE_STYLES.flatMap((row, r) =>
          row.map((s, c) => (
            <button
              key={s.name}
              type="button"
              role="menuitem"
              className="shape-styles-item"
              title={s.name}
              aria-label={s.name}
              onClick={() => {
                env.c.objects.update("Shape Style", (o) => {
                  if (!isChart(o) && o.kind === "shape" && !o.connector && !o.ink) applyShapeStyle(o, s)
                })
                close()
              }}
            >
              <svg
                width="56"
                height="36"
                viewBox="0 0 56 36"
                aria-hidden
                dangerouslySetInnerHTML={{ __html: styleTile(env, s, `ssfx${r}-${c}`) }}
              />
            </button>
          )),
        )}
      </div>
    </div>
  ),
}

/** A colour menu for shapes: picks a theme-aware colour (null: none). */
function colorMenu(noneLabel: string, pick: (env: RibbonEnv, color: ShapeColor | null) => void): MenuEntry {
  return {
    panel: (env, close) => (
      <ColorPanel
        noneLabel={noneLabel}
        theme={env.c.wb.theme}
        onPick={(hex, ref) => {
          pick(env, hex ? shapeColorOf(hex, ref) : null)
          close()
        }}
      />
    ),
  }
}

const WEIGHTS = [0.25, 0.5, 0.75, 1, 1.5, 2.25, 3, 4.5, 6]
const DASH_NAMES: [DashStyle, string][] = [
  ["solid", "Solid"],
  ["sysDot", "Round Dot"],
  ["sysDash", "Square Dot"],
  ["dash", "Dash"],
  ["dashDot", "Dash Dot"],
  ["lgDash", "Long Dash"],
]

/** Insert › Illustrations and Text, Shape Format, Picture Format, Page Layout › Arrange. */
export function objectCommands(): Command[] {
  const commands: Command[] = [
    // ── Insert ──
    {
      id: "ShapesInsertGallery",
      label: "Shapes",
      description:
        "Insert ready-made shapes, such as circles, squares and arrows. Click a shape, then drag on the sheet.",
      keytip: "SH",
      menu: [shapesPanel],
    },
    {
      id: "TextBoxInsertExcel",
      label: "Text Box",
      description: "Draw a text box anywhere on the sheet: click or drag where it goes, then type.",
      keytip: "X",
      run: (env) => env.c.objects.startShape("rect", true),
      pressed: (env) => !!env.c.objects.pending?.textBox,
    },
    {
      id: "WordArtInsertGallery",
      label: "WordArt",
      description: "Add a touch of artistry with decorative text in a text box.",
      keytip: "W",
      menu: [wordArtPanel],
    },
    {
      id: "FlyoutAnchorInsertPictures",
      label: "Pictures",
      description: "Insert pictures from this device: photos, logos and other images.",
      keytip: "P",
      run: (env) => {
        void pickPicture().then((p) => {
          if (!p) return
          env.c.objects.insertPicture(p.src, { w: p.w, h: p.h })
          env.c.shell.showTab?.("TabPictureToolsFormat")
        })
      },
    },
    {
      id: "IconInsertFromFile",
      label: "Icons",
      description:
        "Insert an icon to show something with a picture. The icons come with Cells, so nothing goes online.",
      keytip: "Y1",
      run: (env) => env.dialog("icons"),
    },

    // ── Shape Format › Shape Styles ──
    {
      id: "ShapeStylesGallery",
      label: "Shape Styles",
      description:
        "Choose a ready-made look for the shape: its fill, outline and text colour, and in the lower rows the theme's effects.",
      keytip: "SS",
      enabled: hasShape,
      menu: [shapeStylesPanel],
    },
    {
      id: "ShapeFillColorPicker",
      label: "Shape Fill",
      description: "Fill the selected shape with a solid colour.",
      keytip: "SI",
      enabled: hasShape,
      menu: [
        colorMenu("No Fill", (env, color) =>
          env.c.objects.update("Shape Fill", (o) => {
            if (isChart(o) || o.kind !== "shape" || o.connector) return
            o.fill = color ? { kind: "solid", color } : { kind: "none" }
          }),
        ),
      ],
    },
    {
      id: "ShapeOutlineColorPicker",
      label: "Shape Outline",
      description: "Choose the colour, width and line style of the selected shape's outline.",
      keytip: "SO",
      enabled: hasDrawing,
      menu: [
        colorMenu("No Outline", (env, color) =>
          env.c.objects.update("Shape Outline", (o) => {
            if (isChart(o) || o.kind === "group" || o.kind === "slicer" || o.kind === "timeline") return
            const line = o.kind === "shape" ? o.line : (o.line ?? { color: null, width: 0.75 })
            o.line = { ...line, color, width: line.width || 0.75 }
          }),
        ),
        "-",
        { heading: "Weight" },
        ...WEIGHTS.map((w) => `ShapeOutlineWeight:${w}`),
        { heading: "Dashes" },
        ...DASH_NAMES.map(([d]) => `ShapeOutlineDash:${d}`),
      ],
    },
    ...WEIGHTS.map((w): Command => ({
      id: `ShapeOutlineWeight:${w}`,
      label: `${w} pt`,
      run: (env) =>
        env.c.objects.update("Outline Weight", (o) => {
          if (isChart(o) || o.kind === "group" || o.kind === "slicer" || o.kind === "timeline") return
          const line = o.kind === "shape" ? o.line : (o.line ?? { color: { theme: "dk1" }, width: w })
          o.line = { ...line, width: w, color: line.color ?? { theme: "dk1" } }
        }),
      pressed: (env) => shapes(env)[0]?.line.width === w,
    })),
    ...DASH_NAMES.map(([d, name]): Command => ({
      id: `ShapeOutlineDash:${d}`,
      label: name,
      run: (env) =>
        env.c.objects.update("Outline Dash", (o) => {
          if (isChart(o) || o.kind !== "shape") return
          o.line = { ...o.line, dash: d }
        }),
      pressed: (env) => (shapes(env)[0]?.line.dash ?? "solid") === d,
    })),
    {
      id: "ShapeEffectsMenu",
      label: "Shape Effects",
      description: "Give the selected shape a shadow.",
      keytip: "SE",
      enabled: hasShape,
      menu: ["ShapeShadowToggle"],
    },
    {
      id: "ShapeShadowToggle",
      label: "Shadow",
      run: (env) => {
        const on = !shapes(env).every((s) => s.shadow)
        env.c.objects.update("Shape Effects", (o) => {
          if (isChart(o) || o.kind !== "shape") return
          if (on) o.shadow = true
          else delete o.shadow
        })
      },
      pressed: (env) => shapes(env).length > 0 && shapes(env).every((s) => s.shadow),
    },

    // ── Shape Format › WordArt Styles ──
    {
      id: "TextFillColorPicker",
      label: "Text Fill",
      description: "Colour the text in the selected shape.",
      keytip: "TI",
      enabled: hasShape,
      menu: [
        colorMenu("No Fill", (env, color) =>
          env.c.objects.update("Text Fill", (o) => {
            if (isChart(o) || o.kind !== "shape" || !o.text) return
            if (o.wordArt) o.wordArt = { ...o.wordArt, fill: color ? { kind: "solid", color } : { kind: "none" } }
            for (const p of o.text.paragraphs) for (const r of p.runs) r.color = color ?? undefined
          }),
        ),
      ],
    },
    {
      id: "TextOutlineColorPicker",
      label: "Text Outline",
      description: "Outline the letters of the selected shape's text.",
      keytip: "TO",
      enabled: hasShape,
      menu: [
        colorMenu("No Outline", (env, color) =>
          env.c.objects.update("Text Outline", (o) => {
            if (isChart(o) || o.kind !== "shape") return
            const effects = { ...o.wordArt?.effects }
            if (color) effects.outline = { color, width: 0.75 }
            else delete effects.outline
            o.wordArt = { ...o.wordArt, effects }
          }),
        ),
      ],
    },
    {
      id: "TextEffectsMenu",
      label: "Text Effects",
      description: "Give the selected shape's text a shadow, a glow or a reflection.",
      keytip: "TE",
      enabled: hasShape,
      menu: ["TextEffectShadow", "TextEffectGlow", "TextEffectReflection"],
    },
    ...(["shadow", "glow", "reflection"] as const).map((effect): Command => ({
      id: `TextEffect${effect[0].toUpperCase()}${effect.slice(1)}`,
      label: effect[0].toUpperCase() + effect.slice(1),
      run: (env) => {
        const on = !shapes(env).every((s) => !!s.wordArt?.effects?.[effect])
        env.c.objects.update("Text Effects", (o) => {
          if (isChart(o) || o.kind !== "shape") return
          const effects = { ...o.wordArt?.effects }
          if (!on) delete effects[effect]
          else if (effect === "glow") effects.glow = { theme: "accent1", alpha: 0.6 }
          else effects[effect] = true
          o.wordArt = { ...o.wordArt, effects }
        })
      },
      pressed: (env) => shapes(env).length > 0 && shapes(env).every((s) => !!s.wordArt?.effects?.[effect]),
    })),

    // ── Picture Format ──
    {
      id: "PictureBorderColorPicker",
      label: "Picture Border",
      description: "Choose the colour of the selected picture's border.",
      keytip: "PB",
      enabled: (env) => !!env.c.objects.picture(),
      menu: [
        colorMenu("No Outline", (env, color) =>
          env.c.objects.update("Picture Border", (o) => {
            if (isChart(o) || o.kind !== "picture") return
            if (color) o.line = { color, width: o.line?.width ?? 0.75 }
            else delete o.line
          }),
        ),
      ],
    },
    {
      id: "PictureReset",
      label: "Reset Picture",
      description: "Discard the size and crop changes made to this picture.",
      keytip: "Q",
      enabled: (env) => !!env.c.objects.picture()?.natural,
      run: (env) => {
        const p = env.c.objects.picture()
        if (!p?.natural) return
        const box = env.c.objects.box(p)
        env.c.objects.update("Reset Picture", (o) => {
          if (isChart(o) || o.kind !== "picture") return
          delete o.crop
          delete o.rotation
        })
        env.c.objects.place(new Map([[p.id, { ...box, w: p.natural.w, h: p.natural.h }]]), "Reset Picture")
      },
    },

    // ── Size ──
    sizeField("ShapeHeight", "Height", "h"),
    sizeField("ShapeWidth", "Width", "w"),

    // ── Arrange (Shape Format, Picture Format and Page Layout) ──
    {
      id: "ObjectBringForwardMenu",
      label: "Bring Forward",
      description: "Bring the selected object forward so that it is hidden behind fewer objects.",
      keytip: "AF",
      enabled: hasObjects,
      run: (env) => env.c.objects.order("forward"),
      menu: ["ObjectBringForward", "ObjectBringToFront"],
    },
    {
      id: "ObjectBringForward",
      label: "Bring Forward",
      run: (env) => env.c.objects.order("forward"),
      enabled: hasObjects,
    },
    {
      id: "ObjectBringToFront",
      label: "Bring to Front",
      run: (env) => env.c.objects.order("front"),
      enabled: hasObjects,
    },
    {
      id: "ObjectSendBackwardMenu",
      label: "Send Backward",
      description: "Send the selected object backward so that it is hidden behind more objects.",
      keytip: "AE",
      enabled: hasObjects,
      run: (env) => env.c.objects.order("backward"),
      menu: ["ObjectSendBackward", "ObjectSendToBack"],
    },
    {
      id: "ObjectSendBackward",
      label: "Send Backward",
      run: (env) => env.c.objects.order("backward"),
      enabled: hasObjects,
    },
    { id: "ObjectSendToBack", label: "Send to Back", run: (env) => env.c.objects.order("back"), enabled: hasObjects },
    {
      id: "SelectionPane",
      label: "Selection Pane",
      description:
        "See a list of all your objects. This makes it easier to select objects, change their order, or change their visibility.",
      keytip: "AP",
      run: (env) => env.c.objects.setPane(),
      pressed: (env) => env.c.objects.paneOpen,
    },
    {
      id: "SelectObjects",
      label: "Select Objects",
      description: "Select every object on the sheet: shapes, pictures, text boxes and charts.",
      run: (env) => env.c.objects.selectAll(),
      enabled: (env) => env.c.sheet.objects.length > 0,
    },
    {
      id: "ObjectAlignMenu",
      label: "Align",
      description: "Change the placement of the selected objects on the sheet: line them up, or space them evenly.",
      keytip: "AA",
      enabled: hasObjects,
      menu: [
        "ObjectsAlignLeft",
        "ObjectsAlignCenterHorizontal",
        "ObjectsAlignRight",
        "-",
        "ObjectsAlignTop",
        "ObjectsAlignMiddleVertical",
        "ObjectsAlignBottom",
        "-",
        "AlignDistributeHorizontallyClassic",
        "AlignDistributeVerticallyClassic",
      ],
    },
    align("ObjectsAlignLeft", "Align Left", "left"),
    align("ObjectsAlignCenterHorizontal", "Align Center", "center"),
    align("ObjectsAlignRight", "Align Right", "right"),
    align("ObjectsAlignTop", "Align Top", "top"),
    align("ObjectsAlignMiddleVertical", "Align Middle", "middle"),
    align("ObjectsAlignBottom", "Align Bottom", "bottom"),
    {
      ...align("AlignDistributeHorizontallyClassic", "Distribute Horizontally", "distributeH"),
      enabled: (env) => env.c.objects.selected.length >= 3,
    },
    {
      ...align("AlignDistributeVerticallyClassic", "Distribute Vertically", "distributeV"),
      enabled: (env) => env.c.objects.selected.length >= 3,
    },
    {
      id: "ObjectsGroupMenu",
      label: "Group",
      description: "Join objects together to move and format them as if they were a single object.",
      keytip: "AG",
      enabled: hasDrawing,
      menu: ["ObjectsGroup", "ObjectsUngroup"],
    },
    {
      id: "ObjectsGroup",
      label: "Group",
      run: (env) => env.c.objects.group(),
      // Slicers stay on their own (Cells can't keep one inside a group).
      enabled: (env) =>
        env.c.objects.drawings().length >= 2 &&
        !env.c.objects.drawings().some((o) => o.kind === "slicer" || o.kind === "timeline"),
    },
    {
      id: "ObjectsUngroup",
      label: "Ungroup",
      run: (env) => env.c.objects.ungroup(),
      // SmartArt comes apart only by Convert to Shapes.
      enabled: (env) => env.c.objects.drawings().some((o) => o.kind === "group" && !o.smartArt),
    },
    {
      id: "ObjectRotateGallery",
      label: "Rotate",
      description: "Rotate or flip the selected object.",
      keytip: "AY",
      // Slicers don't turn, as in Excel.
      enabled: (env) =>
        hasDrawing(env) && !env.c.objects.drawings().some((o) => o.kind === "slicer" || o.kind === "timeline"),
      menu: ["ObjectRotateRight90", "ObjectRotateLeft90", "ObjectFlipVertical", "ObjectFlipHorizontal"],
    },
    {
      id: "ObjectRotateRight90",
      label: "Rotate Right 90°",
      run: (env) => env.c.objects.rotate(90),
      enabled: hasDrawing,
    },
    {
      id: "ObjectRotateLeft90",
      label: "Rotate Left 90°",
      run: (env) => env.c.objects.rotate(-90),
      enabled: hasDrawing,
    },
    { id: "ObjectFlipVertical", label: "Flip Vertical", run: (env) => env.c.objects.flip("v"), enabled: hasDrawing },
    {
      id: "ObjectFlipHorizontal",
      label: "Flip Horizontal",
      run: (env) => env.c.objects.flip("h"),
      enabled: hasDrawing,
    },
    {
      id: "ObjectAlternativeText",
      label: "Alt Text",
      description: "Describe the object for people who use a screen reader, or mark it as decorative.",
      keytip: "AT",
      enabled: (env) => env.c.objects.active() !== null,
      run: (env) => {
        const o = env.c.objects.active()
        if (o) env.c.shell.openDialog?.(`altText:${o.id}`)
      },
    },
  ]
  return commands
}

function align(id: string, label: string, to: Parameters<RibbonEnv["c"]["objects"]["align"]>[0]): Command {
  return { id, label, run: (env) => env.c.objects.align(to), enabled: hasObjects }
}

/** Shape Format › Size: the selected objects' height or width, in centimetres or inches as the region uses. */
function sizeField(id: string, label: string, side: "w" | "h"): Command {
  const metric = () => !/^en-US/.test(navigator.language ?? "en-US")
  const unit = () => (metric() ? "cm" : '"')
  const per = () => (metric() ? 96 / 2.54 : 96)
  return {
    id,
    label,
    description: `The ${label.toLowerCase()} of the selected object.`,
    keytip: side === "h" ? "HE" : "WI",
    enabled: hasObjects,
    field: {
      width: 64,
      value: (env) => {
        const o = env.c.objects.active()
        if (!o) return ""
        const px = env.c.objects.box(o)[side]
        return `${(px / per()).toFixed(2).replace(/\.?0+$/, "")}${unit()}`
      },
      commit: (env, text) => {
        const n = parseFloat(text.replace(",", "."))
        if (!Number.isFinite(n) || n <= 0) return
        const cm = /cm/.test(text) || (!/["in]/.test(text) && metric())
        const px = cm ? (n * 96) / 2.54 : n * 96
        const picture = env.c.objects.picture()
        env.c.objects.resize(side === "w" ? px : null, side === "h" ? px : null, !!picture)
      },
    },
  }
}
