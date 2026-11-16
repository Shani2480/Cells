#!/usr/bin/env node
// Frame times inside Cells' web view on an Android emulator or device (debug builds only, which allow the web view's
// DevTools): records the gap between drawn frames for a few seconds while something else swipes the sheet, then prints
// percentiles and the share of frames over budget.
//
//   node scripts/android-frames.mjs <port> <seconds>
//
// with `adb forward tcp:<port> localabstract:webview_devtools_remote_<pid>` set up first (android-smoke.sh does).

const [port = "9333", seconds = "6"] = process.argv.slice(2)
const pages = await (await fetch(`http://127.0.0.1:${port}/json`)).json()
const page = pages.find((p) => p.type === "page" && p.webSocketDebuggerUrl)
if (!page) {
  console.error("No page in the web view.")
  process.exit(1)
}

const socket = new WebSocket(page.webSocketDebuggerUrl)
await new Promise((resolve, reject) => {
  socket.onopen = resolve
  socket.onerror = reject
})
let id = 0
const waiting = new Map()
socket.onmessage = (event) => {
  const message = JSON.parse(event.data)
  waiting.get(message.id)?.(message)
  waiting.delete(message.id)
}
const evaluate = (expression) =>
  new Promise((resolve) => {
    const n = ++id
    waiting.set(n, (m) => resolve(m.result?.result?.value))
    socket.send(
      JSON.stringify({
        id: n,
        method: "Runtime.evaluate",
        params: { expression, awaitPromise: true, returnByValue: true },
      }),
    )
  })

// Only frames that draw count: an idle page makes no frames, so gaps over 250 ms are pauses between swipes.
const gaps = await evaluate(`new Promise((done) => {
  const gaps = []
  let last = performance.now()
  const end = last + ${Number(seconds) * 1000}
  const frame = (now) => {
    gaps.push(now - last)
    last = now
    if (now < end) requestAnimationFrame(frame)
    else done(gaps.filter((g) => g < 250))
  }
  requestAnimationFrame(frame)
})`)
socket.close()

const sorted = [...gaps].sort((a, b) => a - b)
const at = (q) => sorted[Math.min(sorted.length - 1, Math.floor(q * sorted.length))] ?? 0
const over = (ms) => (sorted.length ? (100 * sorted.filter((g) => g > ms).length) / sorted.length : 0)
// A frame later than one and a half budgets is a dropped one (timer jitter puts many a hair over 16.7 ms).
console.log(
  `    frames ${sorted.length}, median ${at(0.5).toFixed(1)} ms, 90th ${at(0.9).toFixed(1)} ms, 99th ${at(0.99).toFixed(1)} ms; ` +
    `dropped (over 25 ms): ${over(25).toFixed(1)}%`,
)
