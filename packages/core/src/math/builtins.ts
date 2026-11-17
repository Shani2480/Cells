/**
 * Insert › Equation's gallery and the Equation tab's Structures and Symbols, as linear text (□ marks a place to
 * type). The gallery holds Office's built-in equations.
 */

export const BUILTIN_EQUATIONS: { name: string; linear: string }[] = [
  { name: "Area of Circle", linear: "A=πr^2" },
  { name: "Binomial Theorem", linear: "(x+a)^n=∑_(k=0)^n▒〖(n¦k) x^k a^(n−k)〗" },
  { name: "Expansion of a Sum", linear: "(1+x)^n=1+nx/1!+(n(n−1) x^2)/2!+…" },
  { name: "Fourier Series", linear: "f(x)=a_0+∑_(n=1)^∞▒(a_n cos⁡〖nπx/L〗+b_n sin⁡〖nπx/L〗)" },
  { name: "Pythagorean Theorem", linear: "a^2+b^2=c^2" },
  { name: "Quadratic Formula", linear: "x=(−b±√(b^2−4ac))/2a" },
  { name: "Taylor Expansion", linear: "e^x=1+x/1!+x^2/2!+x^3/3!+…,  −∞<x<∞" },
  { name: "Trig Identity 1", linear: "sin⁡α±sin⁡β=2 sin⁡〖1/2 (α±β)〗 cos⁡〖1/2 (α∓β)〗" },
  { name: "Trig Identity 2", linear: "cos⁡α+cos⁡β=2 cos⁡〖1/2 (α+β)〗 cos⁡〖1/2 (α−β)〗" },
]

export type MathStructure = { name: string; linear: string }

/** The Equation tab's Structures group: each menu's templates. */
export const MATH_STRUCTURES: { id: string; label: string; items: MathStructure[] }[] = [
  {
    id: "Fraction",
    label: "Fraction",
    items: [
      { name: "Stacked Fraction", linear: "□/□" },
      { name: "Linear Fraction", linear: "□∕□" },
      { name: "dy over dx", linear: "dy/dx" },
      { name: "Partial Differential", linear: "∂y/∂x" },
      { name: "Pi Over 2", linear: "π/2" },
    ],
  },
  {
    id: "Script",
    label: "Script",
    items: [
      { name: "Superscript", linear: "□^□" },
      { name: "Subscript", linear: "□_□" },
      { name: "Subscript-Superscript", linear: "□_□^□" },
      { name: "Subscript x squared", linear: "x_1^2" },
      { name: "e to the minus i omega t", linear: "e^(−iωt)" },
    ],
  },
  {
    id: "Radical",
    label: "Radical",
    items: [
      { name: "Square Root", linear: "√□" },
      { name: "Radical with Degree", linear: "√(□&□)" },
      { name: "Cube Root", linear: "∛□" },
      { name: "Square Root of a squared plus b squared", linear: "√(a^2+b^2)" },
    ],
  },
  {
    id: "Integral",
    label: "Integral",
    items: [
      { name: "Integral", linear: "∫▒□" },
      { name: "Integral with Limits", linear: "∫_□^□▒□" },
      { name: "Double Integral", linear: "∬▒□" },
      { name: "Contour Integral", linear: "∮▒□" },
      { name: "Differential x", linear: "∫_a^b▒f(x)dx" },
    ],
  },
  {
    id: "LargeOperator",
    label: "Large Operator",
    items: [
      { name: "Summation", linear: "∑▒□" },
      { name: "Summation with Limits", linear: "∑_□^□▒□" },
      { name: "Product with Limits", linear: "∏_□^□▒□" },
      { name: "Union with Limits", linear: "⋃_□^□▒□" },
      { name: "Summation over k of n choose k", linear: "∑_(k=0)^n▒(n¦k)" },
    ],
  },
  {
    id: "Bracket",
    label: "Bracket",
    items: [
      { name: "Parentheses", linear: "(□)" },
      { name: "Brackets", linear: "[□]" },
      { name: "Braces", linear: "{□}" },
      { name: "Absolute Value", linear: "|□|" },
      { name: "Norm", linear: "‖□‖" },
      { name: "Angle Brackets", linear: "⟨□⟩" },
      { name: "Floor", linear: "⌊□⌋" },
      { name: "Ceiling", linear: "⌈□⌉" },
      { name: "Cases", linear: "{■(□&□@□&□)┤" },
      { name: "Binomial Coefficient", linear: "(□¦□)" },
    ],
  },
  {
    id: "Function",
    label: "Function",
    items: [
      { name: "Sine Function", linear: "sin⁡□" },
      { name: "Cosine Function", linear: "cos⁡□" },
      { name: "Tangent Function", linear: "tan⁡□" },
      { name: "Natural Logarithm", linear: "ln⁡□" },
      { name: "Logarithm with Base", linear: "log_□⁡□" },
      { name: "Exponential", linear: "exp⁡□" },
    ],
  },
  {
    id: "Accent",
    label: "Accent",
    items: [
      { name: "Hat", linear: "□̂" },
      { name: "Tilde", linear: "□̃" },
      { name: "Dot", linear: "□̇" },
      { name: "Double Dot", linear: "□̈" },
      { name: "Bar", linear: "□̅" },
      { name: "Vector Arrow", linear: "□⃗" },
    ],
  },
  {
    id: "LimitLog",
    label: "Limit and Log",
    items: [
      { name: "Limit", linear: "lim┬(□→□)⁡□" },
      { name: "Limit as n approaches infinity", linear: "lim┬(n→∞)⁡〖(1+1/n)^n〗" },
      { name: "Maximum", linear: "max┬□⁡□" },
      { name: "Minimum", linear: "min┬□⁡□" },
    ],
  },
  {
    id: "Operator",
    label: "Operator",
    items: [
      { name: "Colon Equal", linear: "≔" },
      { name: "Defined As", linear: "≝" },
      { name: "Right Arrow", linear: "→" },
      { name: "Left Right Arrow", linear: "↔" },
      { name: "Approximately Equal", linear: "≈" },
    ],
  },
  {
    id: "Matrix",
    label: "Matrix",
    items: [
      { name: "2×2 Empty Matrix", linear: "■(□&□@□&□)" },
      { name: "2×2 Matrix with Parentheses", linear: "(■(□&□@□&□))" },
      { name: "2×2 Matrix with Brackets", linear: "[■(□&□@□&□)]" },
      { name: "3×3 Identity Matrix", linear: "(■(1&0&0@0&1&0@0&0&1))" },
      { name: "1×3 Empty Matrix", linear: "■(□&□&□)" },
      { name: "3×1 Empty Matrix", linear: "■(□@□@□)" },
    ],
  },
]

/** The Equation tab's Symbols gallery (Basic Math and Greek letters, as Office lists them). */
export const MATH_SYMBOLS = [
  ..."±∞=≠∼×÷!∝<≪>≫≤≥∓≅≈≡∀∁∂√∛∜∪∩∅%°℉℃∆∇∃∄∈∋←↑→↓↔∴+−¬",
  ..."αβγδεϵζηθϑικλμνξπϖρϱσςτυφϕχψω",
  ..."ΓΔΘΛΞΠΣΥΦΨΩ",
  ..."∗∙⋮⋯⋰⋱ℵℶ∎",
]
