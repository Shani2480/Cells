import type { Command } from "./registry"

/** Insert › Links and Text and Symbols: Link (Ctrl+K), Header & Footer and Symbol. */
export function linkCommands(): Command[] {
  return [
    {
      id: "InsertLinkGallery",
      label: "Link",
      description:
        "Create a link in your document for quick access to webpages and files, a place in this workbook, or an e-mail address.",
      shortcut: "Ctrl+K",
      keytip: "I",
      run: (env) => env.dialog("insertLink"),
    },
    {
      id: "HeaderFooterInsert",
      label: "Header & Footer",
      description:
        "Edit the header or footer of your document. They repeat at the top and bottom of each printed page.",
      keytip: "H",
      // Page Setup's tab, until Page Layout view (where Excel edits them on the page) is built.
      run: (env) => env.dialog("pageSetup:headerFooter"),
    },
    {
      id: "SymbolInsert",
      label: "Symbol",
      description:
        "Add characters that aren't on your keyboard, such as copyright symbols, arrows and accented letters.",
      keytip: "U",
      run: (env) => env.dialog("symbol"),
    },
  ]
}
