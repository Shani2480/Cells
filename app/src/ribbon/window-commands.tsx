import { useEffect, useState } from "react"
import { Check } from "lucide-react"
import { switchEntries, type SwitchEntry } from "@/file/switching"
import { usePlatform } from "@/platform/context"
import type { Command, RibbonEnv } from "./registry"

/**
 * View › Window: New Window, Arrange All, Hide / Unhide, View Side by Side (Synchronous Scrolling, Reset Window
 * Position) and Switch Windows. The windows are panels inside Cells' own window (file/windows.ts); Switch Windows
 * also lists the other workbooks open in it.
 */

const windowsOf = (env: RibbonEnv) => env.files?.doc.windows

/** Switch Windows' list: every window of every workbook, the one shown ticked (see file/switching.ts). */
function SwitchWindowsList({ env, close }: { env: RibbonEnv; close: () => void }) {
  const platform = usePlatform()
  const [list, setList] = useState<SwitchEntry[] | null>(null)
  useEffect(() => {
    let live = true
    if (env.files) void switchEntries(env.files, platform.window).then((got) => live && setList(got))
    return () => {
      live = false
    }
  }, [env.files, platform])
  if (!list) return null
  return (
    <div className="pl-list win-switch" role="menu" aria-label="Switch Windows">
      {list.map((e, i) => (
        <button
          key={e.key}
          type="button"
          role="menuitemradio"
          aria-checked={e.current}
          className={`pl-item win-switch-item${e.current ? " is-current" : ""}`}
          onClick={() => {
            close()
            e.go()
          }}
        >
          <span className="win-switch-check">{e.current ? <Check size={14} /> : null}</span>
          <span>
            <u>{i + 1}</u> {e.title}
          </span>
        </button>
      ))}
    </div>
  )
}

export function windowCommands(): Command[] {
  return [
    {
      id: "WindowNew",
      label: "New Window",
      description: "Open a second window for the workbook, so you can work in different places at the same time.",
      keytip: "N",
      enabled: (env) => windowsOf(env) !== undefined,
      run: (env) => {
        windowsOf(env)?.newWindow()
      },
    },
    {
      id: "WindowsArrangeAll",
      label: "Arrange All",
      description: "Stack the workbook's windows so you can see them all at once.",
      keytip: "A",
      enabled: (env) => windowsOf(env) !== undefined,
      run: (env) => env.dialog("arrangeWindows"),
    },
    {
      id: "WindowHide",
      label: "Hide",
      description: "Hide the current window so it can't be seen. To bring it back, click Unhide.",
      keytip: "H",
      // Cells shows one workbook at a time: hiding its last window would leave nothing to work in.
      enabled: (env) => (windowsOf(env)?.visible.length ?? 0) > 1,
      run: (env) => windowsOf(env)?.hide(),
    },
    {
      id: "WindowUnhide",
      label: "Unhide...",
      description: "Show windows that were hidden with the Hide Window feature.",
      keytip: "U",
      enabled: (env) => (windowsOf(env)?.hidden.length ?? 0) > 0,
      run: (env) => env.dialog("unhideWindow"),
    },
    {
      id: "ViewSideBySide",
      label: "View Side by Side",
      description: "View two windows one above the other to compare their contents.",
      keytip: "B",
      enabled: (env) => (windowsOf(env)?.all.length ?? 0) > 1,
      pressed: (env) => Boolean(windowsOf(env)?.sideBySide),
      run: (env) => {
        const w = windowsOf(env)
        if (!w) return
        if (w.sideBySide) return w.endSideBySide()
        const others = w.all.filter((c) => c !== w.active)
        if (others.length === 1) w.startSideBySide(others[0]!)
        else env.dialog("compareSideBySide")
      },
    },
    {
      id: "WindowSideBySideSynchronousScrolling",
      label: "Synchronous Scrolling",
      description: "Scroll two windows at the same time. Turn on View Side by Side to use it.",
      keytip: "YS",
      enabled: (env) => Boolean(windowsOf(env)?.sideBySide),
      pressed: (env) => Boolean(windowsOf(env)?.sideBySide && windowsOf(env)?.syncScroll),
      run: (env) => windowsOf(env)?.toggleSyncScroll(),
    },
    {
      id: "WindowResetPosition",
      label: "Reset Window Position",
      description: "Give the two windows compared side by side an equal share of the screen.",
      keytip: "T",
      enabled: (env) => Boolean(windowsOf(env)?.sideBySide),
      run: (env) => windowsOf(env)?.resetPosition(),
    },
    {
      id: "WindowSwitchWindowsMenuExcel",
      label: "Switch Windows",
      description: "Switch to another window: another of this workbook's, or another open workbook.",
      keytip: "W",
      enabled: (env) => windowsOf(env) !== undefined,
      menu: [
        {
          panel: (env, close) => <SwitchWindowsList env={env} close={close} />,
        },
      ],
    },
  ]
}
