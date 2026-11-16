#!/usr/bin/env bash
# Android smoke test (docs/release.md): on one emulator, install the debug build, open a workbook as "Open with"
# does, type into a cell with hardware keys, save with Ctrl+S, pull the file back and check the entry is in it. Then
# swipe the sheet and report frame times, and take light and dark screenshots.
#
#   scripts/android-smoke.sh <emulator serial> <phone|tablet> [screenshot folder]
#
# Emulators only: a real device's status bar shows its owner's notifications. Build first with
#   npm run tauri android build -- --debug --target x86_64 --apk
set -euo pipefail

serial=${1:?emulator serial, e.g. emulator-5554}
kind=${2:?phone or tablet}
shots=${3:-}
case "$serial" in emulator-*) ;; *) echo "Only emulators: $serial is a device." >&2; exit 2 ;; esac

root=$(cd "$(dirname "$0")/.." && pwd)
adb="${ANDROID_HOME:-$HOME/Android/Sdk}/platform-tools/adb -s $serial"
apk=$root/src-tauri/gen/android/app/build/outputs/apk/universal/debug/app-universal-debug.apk
book=$root/packages/io/test/fixtures/excel/budget.xlsx
remote="/sdcard/Download/cells-smoke.xlsx"
mime=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet
work=$(mktemp -d)
trap 'rm -rf "$work"' EXIT
step() { printf '• %s\n' "$*"; }
# Taps the first thing on screen with this text (the page's controls are in the accessibility tree too).
tap_text() {
  $adb shell uiautomator dump /sdcard/cells-ui.xml >/dev/null 2>&1
  local xy
  xy=$($adb exec-out cat /sdcard/cells-ui.xml | python3 -I -c '
import re, sys
text = sys.argv[1]
for m in re.finditer(r"<node [^>]*?(?:text|content-desc)=\"([^\"]*)\"[^>]*?bounds=\"\[(\d+),(\d+)\]\[(\d+),(\d+)\]\"", sys.stdin.read()):
    if m.group(1) == text:
        a, b, c, d = map(int, m.groups()[1:])
        print((a + c) // 2, (b + d) // 2)
        break
' "$1")
  $adb shell rm -f /sdcard/cells-ui.xml
  [ -n "$xy" ] || return 1
  $adb shell input tap $xy
}

[ "$($adb shell getprop sys.boot_completed | tr -d '\r')" = 1 ] || { echo "$serial hasn't finished booting." >&2; exit 2; }
size=$($adb shell wm size | sed -n 's/.*: //p' | tr -d '\r')
step "$serial ($kind, $size)"

step "Install"
$adb install -r "$apk" >/dev/null
$adb shell pm clear app.cells >/dev/null
# A clean status bar for the screenshots.
$adb shell settings put global sysui_demo_allowed 1
$adb shell am broadcast -a com.android.systemui.demo -e command enter >/dev/null
$adb shell am broadcast -a com.android.systemui.demo -e command clock -e hhmm 0941 >/dev/null
$adb shell am broadcast -a com.android.systemui.demo -e command notifications -e visible false >/dev/null
$adb shell cmd uimode night no >/dev/null

step "Open with: $(basename "$book")"
$adb push "$book" "$remote" >/dev/null
$adb shell am broadcast -a android.intent.action.MEDIA_SCANNER_SCAN_FILE -d "file://$remote" >/dev/null
sleep 2
id=$($adb shell content query --uri content://media/external/file --projection _id \
  --where "\"_data='/storage/emulated/0/Download/cells-smoke.xlsx'\"" | sed -n 's/.*_id=\([0-9]*\).*/\1/p' | head -1)
[ -n "$id" ] || { echo "The media store doesn't list the workbook." >&2; exit 1; }
uri=content://media/external/file/$id
# The first launch: past the setup (Skip keeps the defaults), then the file.
$adb shell monkey -p app.cells -c android.intent.category.LAUNCHER 1 >/dev/null 2>&1
skipped=
for _ in $(seq 1 15); do
  sleep 1
  if tap_text Skip; then skipped=1; break; fi
done
[ -n "$skipped" ] || { echo "The first-run setup didn't show." >&2; exit 1; }
sleep 1
$adb shell am start -a android.intent.action.VIEW -d "$uri" -t $mime \
  --grant-read-uri-permission --grant-write-uri-permission -n app.cells/.MainActivity >/dev/null
sleep 6
title=$($adb shell dumpsys activity top | grep -o 'cells-smoke' | head -1 || true)

step "Edit D1 and save with Ctrl+S"
stamp="smoke-$(date +%s)"
read -r w h <<<"$(echo "$size" | tr 'x' ' ')"
# A tap on the sheet gives it the keys, as a person's first touch does.
$adb shell input tap $((w / 2)) $((h / 2))
sleep 1
$adb shell input keycombination KEYCODE_CTRL_LEFT KEYCODE_MOVE_HOME
for _ in 1 2 3; do $adb shell input keyevent KEYCODE_DPAD_RIGHT; done
$adb shell input text "$stamp"
$adb shell input keyevent KEYCODE_ENTER
sleep 1
$adb shell input keycombination KEYCODE_CTRL_LEFT KEYCODE_S
sleep 4
$adb pull "$remote" "$work/back.xlsx" >/dev/null
if unzip -p "$work/back.xlsx" xl/sharedStrings.xml 2>/dev/null | grep -q "$stamp" \
  && unzip -p "$work/back.xlsx" xl/sharedStrings.xml | grep -q "Rent"; then
  step "Saved back: the new entry and the workbook's own text are in the file"
else
  echo "FAIL: the saved workbook doesn't hold \"$stamp\" (title seen: ${title:-none})." >&2
  exit 1
fi

# Typing brought up the on-screen keyboard; put it away, or the swipes below would type on it.
hide_keyboard() {
  if $adb shell dumpsys input_method | grep -q "mInputShown=true"; then
    $adb shell input keyevent KEYCODE_BACK
    sleep 1
  fi
}
hide_keyboard

step "Swipe the sheet (frame times in the web view)"
pid=$($adb shell pidof app.cells | tr -d '\r')
port=9333
$adb forward tcp:$port localabstract:webview_devtools_remote_$pid >/dev/null
node "$root/scripts/android-frames.mjs" $port 7 &
frames=$!
sleep 1
for _ in 1 2 3 4 5; do
  $adb shell input swipe $((w / 2)) $((h * 3 / 5)) $((w / 2)) $((h / 3)) 150
  sleep 0.6
done
wait $frames || echo "    (no frame times: the web view's DevTools didn't answer)"
$adb forward --remove tcp:$port >/dev/null || true

if [ -n "$shots" ]; then
  mkdir -p "$shots"
  # The workbook as it came: the test entry undone, the top left in view, no keyboard.
  $adb shell input keycombination KEYCODE_CTRL_LEFT KEYCODE_Z
  $adb shell input keycombination KEYCODE_CTRL_LEFT KEYCODE_MOVE_HOME
  sleep 1
  hide_keyboard
  $adb exec-out screencap -p >"$shots/android-$kind-light.png"
  if [ "$kind" = phone ] && tap_text Home; then
    sleep 1.5
    $adb exec-out screencap -p >"$shots/android-phone-palette.png"
    $adb shell input keyevent KEYCODE_BACK
    sleep 1
  fi
  $adb shell cmd uimode night yes >/dev/null
  sleep 3
  $adb exec-out screencap -p >"$shots/android-$kind-dark.png"
  $adb shell cmd uimode night no >/dev/null
  step "Screenshots in $shots"
fi

step "Back: closes the File view first, then sends Cells to the background (still running)"
on_top() { $adb shell dumpsys activity activities | grep -m1 "topResumedActivity" | grep -q "app.cells"; }
tap_text File || true
sleep 1.5
$adb shell input keyevent KEYCODE_BACK
sleep 1.5
on_top || { echo "FAIL: Back with the File view open left Cells." >&2; exit 1; }
$adb shell input keyevent KEYCODE_BACK
sleep 1.5
if on_top || [ -z "$($adb shell pidof app.cells | tr -d '\r')" ]; then
  echo "FAIL: Back with nothing open should put Cells in the background and keep it running." >&2
  exit 1
fi

$adb shell rm -f "$remote"
$adb shell am broadcast -a com.android.systemui.demo -e command exit >/dev/null
step "OK"
