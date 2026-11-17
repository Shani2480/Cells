import { fireEvent, render, screen } from "@testing-library/react"
import { act } from "react"
import { describe, expect, it } from "vitest"
import { newDocument } from "@/file/document"
import { WindowsArea } from "./WindowsArea"

describe("the windows as panels", () => {
  it("shows one grid maximized, a panel per window when arranged; a click makes a panel the active window", () => {
    const doc = newDocument()
    const first = doc.controller
    doc.windows.newWindow()
    render(
      <WindowsArea
        windows={doc.windows}
        baseTitle={doc.title}
        grid={(c, active) => <div data-testid="grid">{`${doc.windows.title(doc.title, c)} ${active}`}</div>}
      />,
    )
    expect(screen.getAllByTestId("grid").map((g) => g.textContent)).toEqual(["Book1:2 true"])

    act(() => doc.windows.arrange("vertical"))
    expect(screen.getAllByTestId("grid").map((g) => g.textContent)).toEqual(["Book1:1 false", "Book1:2 true"])
    fireEvent.pointerDown(screen.getByRole("region", { name: "Book1:1" }))
    expect(doc.controller).toBe(first)

    fireEvent.click(screen.getByRole("button", { name: "Close Book1:2" }))
    expect(screen.getAllByTestId("grid").map((g) => g.textContent)).toEqual(["Book1 true"])
  })
})
