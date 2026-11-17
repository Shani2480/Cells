import { BUILTIN_EQUATIONS, isEquation, MATH_STRUCTURES, MATH_SYMBOLS, parseLinear } from "@cells/core"
import { layoutEquation } from "@/drawing/math"
import { canvasMeasure } from "@/drawing/render"
import type { Command, MenuEntry, RibbonEnv } from "./registry"

/**
 * Insert › Equation and the Equation tab: Office's built-in equations, structures to fill in and math symbols.
 * Equations are typed in their linear form (𝑥=(−𝑏±√(𝑏^2−4𝑎𝑐))/2𝑎) and build up when typing ends.
 */

/** An equation drawn for a menu or a ribbon picture (`size`: its text in pixels). */
export function MathPicture({ linear, size, label }: { linear: string; size: number; label?: string }) {
  const [box] = layoutEquation([parseLinear(linear)], size, "currentColor", canvasMeasure)
  if (!box) return null
  const pad = 1
  const w = Math.ceil(box.w + pad * 2)
  const h = Math.ceil(box.a + box.d + pad * 2)
  return (
    <svg
      className="math-picture"
      width={w}
      height={h}
      viewBox={`0 0 ${w} ${h}`}
      role={label ? "img" : undefined}
      aria-label={label}
      aria-hidden={label ? undefined : true}
      // The layout's own SVG (built from the equation's text, escaped).
      dangerouslySetInnerHTML={{ __html: `<g transform="translate(${pad} ${pad + box.a})">${box.svg}</g>` }}
    />
  )
}

/** A ribbon picture drawn from a piece of an equation, fitted into the control's square. */
function mathIcon(linear: string, size: 16 | 32) {
  return (
    <span className={`math-icon is-${size}`} aria-hidden>
      <MathPicture linear={linear} size={size === 32 ? 15 : 9} />
    </span>
  )
}

/** Puts `linear` in the selected equation where the caret was (or a new equation, when none is selected). */
export function insertMath(env: RibbonEnv, linear: string) {
  const o = env.c.objects.shape()
  if (isEquation(o)) env.c.objects.insertIntoEquation(o.id, linear)
  else {
    env.c.objects.insertEquation(linear)
    env.c.shell.showTab?.("TabEquationToolsDesign")
  }
}

const galleryPanel: MenuEntry = {
  panel: (env, close) => (
    <div className="equation-gallery" role="menu" aria-label="Built-In equations">
      <div className="equation-gallery-heading">Built-In</div>
      {BUILTIN_EQUATIONS.map((e) => (
        <button
          key={e.name}
          type="button"
          role="menuitem"
          className="equation-gallery-item"
          aria-label={e.name}
          onClick={() => {
            close()
            env.c.objects.insertEquation(e.linear)
            env.c.shell.showTab?.("TabEquationToolsDesign")
          }}
        >
          <span className="equation-gallery-name">{e.name}</span>
          <MathPicture linear={e.linear} size={15} />
        </button>
      ))}
    </div>
  ),
}

const symbolsPanel: MenuEntry = {
  panel: (env, close) => (
    <div className="equation-symbols" role="menu" aria-label="Math symbols">
      {MATH_SYMBOLS.map((s) => (
        <button
          key={s}
          type="button"
          role="menuitem"
          className="equation-symbol"
          title={s}
          onClick={() => {
            close()
            insertMath(env, s)
          }}
        >
          {s}
        </button>
      ))}
    </div>
  ),
}

const structurePanel = (items: { name: string; linear: string }[]): MenuEntry => ({
  panel: (env, close) => (
    <div className="equation-structures" role="menu">
      {items.map((s) => (
        <button
          key={s.name}
          type="button"
          role="menuitem"
          className="equation-structure"
          title={s.name}
          aria-label={s.name}
          onClick={() => {
            close()
            insertMath(env, s.linear)
          }}
        >
          <MathPicture linear={s.linear} size={16} />
        </button>
      ))}
    </div>
  ),
})

/** The picture each Structures menu shows on the ribbon. */
const STRUCTURE_PICTURES: Record<string, string> = {
  Fraction: "x/y",
  Script: "e^x",
  Radical: "√x",
  Integral: "∫_(−x)^x▒",
  LargeOperator: "∑_(i=0)^n▒",
  Bracket: "{()}",
  Function: "sin⁡θ",
  Accent: "ä",
  LimitLog: "lim┬(n→∞)",
  Operator: "∆",
  Matrix: "[■(1&0@0&1)]",
}

const editingEquation = (env: RibbonEnv) => isEquation(env.c.objects.shape())

export function equationCommands(): Command[] {
  return [
    {
      id: "InsertBuildingBlocksEquationsGallery",
      label: "Equation",
      description:
        "Add common mathematical equations, or build your own: type it in its linear form (x^2, a/b, √(x)) and it builds up.",
      keytip: "E",
      icon: mathIcon("π", 16),
      largeIcon: mathIcon("π", 32),
      run: (env) => {
        env.c.objects.insertEquation()
        env.c.shell.showTab?.("TabEquationToolsDesign")
      },
      menu: [galleryPanel, "-", "EquationInsertNew"],
    },
    {
      id: "EquationInsertNew",
      label: "Insert New Equation",
      description: "A new equation to type in its linear form.",
      run: (env) => {
        env.c.objects.insertEquation()
        env.c.shell.showTab?.("TabEquationToolsDesign")
      },
    },
    {
      id: "EquationProfessional",
      label: "Professional",
      description: "Show the equation built up, as it prints.",
      icon: mathIcon("a/b", 16),
      enabled: editingEquation,
      pressed: (env) => env.c.objects.editing === null,
      run: (env) => {
        const o = env.c.objects.shape()
        if (o && env.c.objects.editing === o.id) env.c.objects.finishText?.()
      },
    },
    {
      id: "EquationLinearFormat",
      label: "Linear",
      description: "Show the equation as linear text to edit it (x^2 for x², a/b for a fraction).",
      icon: (
        <span className="math-icon is-16 is-linear" aria-hidden>
          a/b
        </span>
      ),
      enabled: editingEquation,
      pressed: (env) => env.c.objects.editing !== null,
      run: (env) => {
        const o = env.c.objects.shape()
        if (o) env.c.objects.startText(o.id)
      },
    },
    {
      id: "EquationSymbolsGallery",
      label: "Symbols",
      description: "Insert math symbols: operators, arrows, Greek letters.",
      icon: mathIcon("±", 16),
      largeIcon: mathIcon("±∞", 32),
      enabled: editingEquation,
      menu: [symbolsPanel],
    },
    ...MATH_STRUCTURES.map((s): Command => ({
      id: `Equation${s.id}Gallery`,
      label: s.label,
      description: `Insert a ${s.label.toLowerCase()} to fill in.`,
      icon: mathIcon(STRUCTURE_PICTURES[s.id] ?? "□", 16),
      largeIcon: mathIcon(STRUCTURE_PICTURES[s.id] ?? "□", 32),
      enabled: editingEquation,
      menu: [structurePanel(s.items)],
    })),
  ]
}
