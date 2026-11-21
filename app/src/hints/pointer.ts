import { create } from "zustand"

/** A callout Help asked for ("Show me", Show Training): shown at once, whatever the hint settings and timing. */
export type Pointed = {
  /** The feature hint's id: seeing the feature this way dismisses its hint too. */
  id: string
  title: string
  text: string
  /** The control's accessible name. */
  target: string
  /** "2 of 8" during Show Training. */
  step?: string
}

export const usePointer = create<{ pointed: Pointed | null }>(() => ({ pointed: null }))

export function pointAt(pointed: Pointed | null) {
  usePointer.setState({ pointed })
}
