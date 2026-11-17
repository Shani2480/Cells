import type { Command } from "./registry"

/**
 * Review › Comments and Notes, Insert › Comment. The ids of the Notes menu's rows are this app's (Excel's own were
 * not in the reference); see docs/verify-list.md.
 */
export function reviewCommands(): Command[] {
  const newComment: Command = {
    id: "ReviewNewComment",
    label: "New Comment",
    description: "Add a comment about the selection, to start a conversation.",
    shortcut: "Ctrl+Shift+F2",
    run: (env) => env.c.annotations.newComment(),
  }
  return [
    newComment,
    { ...newComment, id: "InsertNewComment", label: "Comment" },
    {
      id: "ReviewDeleteComment",
      label: "Delete",
      description: "Delete the selected comment.",
      run: (env) => env.c.annotations.deleteComments(),
      enabled: (env) => env.c.annotations.selectionHas("comment"),
    },
    {
      id: "ReviewShowAllComments",
      label: "Show Comments",
      description: "Show all the comments on this sheet in a pane beside it.",
      run: (env) => env.c.annotations.togglePane(),
      pressed: (env) => env.c.annotations.paneOpen,
    },
    {
      id: "ReviewPreviousComment",
      label: "Previous",
      description: "Go to the previous comment in the sheet.",
      run: (env) => env.c.annotations.step("comment", -1),
      enabled: (env) => env.c.annotations.hasComments(),
    },
    {
      id: "ReviewNextComment",
      label: "Next",
      description: "Go to the next comment in the sheet.",
      run: (env) => env.c.annotations.step("comment", 1),
      enabled: (env) => env.c.annotations.hasComments(),
    },
    {
      id: "ReviewNoteSplit",
      label: "Notes",
      description: "Add a note, or work with the notes on this sheet.",
      menu: [
        "ReviewNewNote",
        "ReviewPreviousNote",
        "ReviewNextNote",
        "-",
        "ReviewShowOrHideNote",
        "ReviewShowAllNotes",
        "-",
        "ReviewConvertNotesToComments",
      ],
    },
    {
      id: "ReviewNewNote",
      label: "New Note",
      labelFor: (env) => {
        const { row, col } = env.c.selection.active
        return env.c.annotations.noteAt(row, col) ? "Edit Note" : "New Note"
      },
      shortcut: "Shift+F2",
      run: (env) => env.c.annotations.newNote(),
    },
    {
      id: "ReviewPreviousNote",
      label: "Previous Note",
      run: (env) => env.c.annotations.step("note", -1),
      enabled: (env) => env.c.annotations.hasNotes(),
    },
    {
      id: "ReviewNextNote",
      label: "Next Note",
      run: (env) => env.c.annotations.step("note", 1),
      enabled: (env) => env.c.annotations.hasNotes(),
    },
    {
      id: "ReviewShowOrHideNote",
      label: "Show/Hide Note",
      run: (env) => env.c.annotations.toggleNote(),
      enabled: (env) => {
        const { row, col } = env.c.selection.active
        return Boolean(env.c.annotations.noteAt(row, col))
      },
    },
    {
      id: "ReviewShowAllNotes",
      label: "Show All Notes",
      run: (env) => env.c.annotations.toggleShowAll(),
      pressed: (env) => env.c.annotations.showAll,
      enabled: (env) => env.c.annotations.hasNotes() || env.c.annotations.showAll,
    },
    {
      id: "ReviewConvertNotesToComments",
      label: "Convert to Comments",
      run: (env) => env.c.annotations.convertToComments(),
      enabled: (env) => env.c.wb.sheets.some((s) => s.annotations.all().some((a) => a.value.kind === "note")),
    },
    {
      id: "ReviewAllowUsersToEditRangesCoauth",
      label: "Allow Edit Ranges",
      description:
        "Choose ranges of this sheet that people can still edit once it's protected, each with its own password if you like.",
      run: (env) => env.dialog("allowEditRanges"),
      // Excel greys it out while the sheet is protected.
      enabled: (env) => !env.c.sheet.protection,
    },
    {
      id: "StatusWorkbookStatistics",
      label: "Workbook Statistics",
      description:
        "See the number of cells with data, formulas, tables, charts, comments and notes in this sheet and the workbook.",
      run: (env) => env.dialog("statistics"),
    },
    {
      // Excel's control id for Review › Protect Workbook; pressed while the structure is protected.
      id: "ReviewRestrictEditing",
      label: "Protect Workbook...",
      description: "Prevent others from making structural changes to your workbook, such as adding or deleting sheets.",
      pressed: (env) => env.c.sheets.isWorkbookProtected(),
      run: (env) =>
        env.c.sheets.isWorkbookProtected()
          ? env.c.sheets.workbookNeedsPassword()
            ? env.dialog("unprotectWorkbook")
            : env.c.sheets.unprotectWorkbook()
          : env.dialog("protectWorkbook"),
    },
  ]
}
