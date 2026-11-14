/** Joins class names, skipping empty ones. (Inkspire uses clsx + tailwind-merge; Cells' ribbon uses plain classes.) */
export function cn(...parts: (string | false | null | undefined)[]): string {
  return parts.filter(Boolean).join(" ")
}
