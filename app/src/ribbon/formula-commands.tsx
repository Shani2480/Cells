import { FunctionSquare } from "lucide-react"
import { FUNCTION_INFO, functionsIn, type FunctionCategory } from "@cells/core"
import { recentFunctions, rememberFunction } from "@/shell/InsertFunctionDialog"
import type { Command, MenuEntry } from "./registry"

/** One command per function (the Function Library menus list them; Tell Me finds them by name). */
const fnId = (name: string) => `Fn:${name}`

function functionCommands(): Command[] {
  return [...FUNCTION_INFO.values()].map((f) => ({
    id: fnId(f.name),
    label: f.name,
    description: `${f.name}(${f.args.join(", ")}) — ${f.description}`,
    run: (env) => {
      rememberFunction(f.name)
      env.c.insertFunction(f.name)
    },
  }))
}

const listOf = (category: FunctionCategory): MenuEntry[] => [...functionsIn(category).map(fnId), "-", "FunctionWizard"]

const library = (id: string, label: string, description: string, category: FunctionCategory): Command => ({
  id,
  label,
  description,
  menu: listOf(category),
})

/**
 * Formulas › Function Library and Defined Names. Excel's More Functions also has Engineering, Cube and Web, which
 * Cells has no functions for yet; they are left out of the menu.
 */
export function formulaCommands(): Command[] {
  return [
    ...functionCommands(),
    {
      id: "FunctionWizard",
      label: "Insert Function...",
      description:
        "Work with the formula in the current cell. You can easily pick functions to use and get help on how to fill out the input values.",
      shortcut: "Shift+F3",
      keytip: "F",
      largeIcon: <FunctionSquare size={32} strokeWidth={1.25} />,
      run: (env) => env.dialog("insertFunction"),
    },
    {
      id: "FunctionsRecentlyUsedtInsertGallery",
      label: "Recently Used",
      description: "Browse and select from a list of recently used functions.",
      menu: () => [...recentFunctions().map(fnId), "-", "FunctionWizard"],
    },
    library(
      "FunctionsFinancialInsertGallery",
      "Financial",
      "Browse and select from a list of financial functions.",
      "Financial",
    ),
    library(
      "FunctionsLogicalInsertGallery",
      "Logical",
      "Browse and select from a list of logical functions.",
      "Logical",
    ),
    library("FunctionsTextInsertGallery", "Text", "Browse and select from a list of text functions.", "Text"),
    library(
      "FunctionsDateTimeInsertGallery",
      "Date & Time",
      "Browse and select from a list of date and time functions.",
      "Date & Time",
    ),
    library(
      "FunctionsLookupReferenceInsertGallery",
      "Lookup & Reference",
      "Browse and select from a list of lookup and reference functions.",
      "Lookup & Reference",
    ),
    library(
      "FunctionsMathTrigInsertGallery",
      "Math & Trig",
      "Browse and select from a list of math and trigonometry functions.",
      "Math & Trig",
    ),
    {
      id: "FunctionsMoreInsertGallery",
      label: "More Functions",
      description: "Browse and select from lists of statistical, information and compatibility functions.",
      menu: ["FunctionsStatistical", "FunctionsInformation", "FunctionsCompatibility"],
    },
    library("FunctionsStatistical", "Statistical", "Statistical functions.", "Statistical"),
    library("FunctionsInformation", "Information", "Information functions.", "Information"),
    library("FunctionsCompatibility", "Compatibility", "Functions kept for older workbooks.", "Compatibility"),
    {
      id: "NameDefineMenu",
      label: "Define Name",
      description: "Define and apply names.",
      run: (env) => env.dialog("defineName"),
      menu: ["NameDefine", "NameApply"],
    },
    { id: "NameDefine", label: "Define Name...", run: (env) => env.dialog("defineName") },
    { id: "NameApply", label: "Apply Names..." },
    {
      id: "NameUseInFormula",
      label: "Use in Formula",
      description: "Choose a name used in this workbook and insert it into the current formula.",
      menu: (env) => [
        ...env.c.names.all().map((n) => `Name:${n.name}`),
        ...(env.c.names.all().length ? ["-"] : []),
        "NamePaste",
      ],
    },
    { id: "NamePaste", label: "Paste Names..." },
  ]
}

/** "Use in Formula" rows: one command per defined name, made when the menu opens. */
export function nameCommand(name: string): Command {
  return {
    id: `Name:${name}`,
    label: name,
    run: (env) => env.c.insertName(name),
  }
}
