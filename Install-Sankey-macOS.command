#!/bin/bash
# Installs the Sankey add-in for Excel on macOS (Microsoft 365).
# Double-click this file, or run:  bash Install-Sankey-macOS.command
# It copies mac/Sankey.xlam into Excel's Startup folder, so Excel loads it every time it starts.

here="$(cd "$(dirname "$0")" && pwd)"
src="$here/mac/Sankey.xlam"
pause() { echo; read -r -p "Press Enter to close" _; }

if [ ! -f "$src" ]; then
  echo "FAILED: mac/Sankey.xlam is missing."
  echo "It is built once on a Windows PC with Excel: run mac\Build-Mac-Addin.cmd there, then copy the folder here."
  pause; exit 1
fi

if pgrep -x "Microsoft Excel" >/dev/null 2>&1; then
  echo "FAILED: Excel is running. Save your work, quit Excel (Cmd+Q) and run this again."
  pause; exit 1
fi

container="$HOME/Library/Group Containers/UBF8T346G9.Office"
uc=""
for d in "User Content.localized" "User Content"; do
  [ -d "$container/$d" ] && uc="$container/$d" && break
done
if [ -z "$uc" ]; then
  echo "FAILED: Excel's data folder was not found ($container)."
  echo "Open Excel once, quit it, and run this again. (If macOS asks whether Terminal may access data from other apps, choose Allow.)"
  pause; exit 1
fi

dest="$uc/Startup.localized/Excel"
[ -d "$uc/Startup.localized" ] || dest="$uc/Startup/Excel"
mkdir -p "$dest" || { echo "FAILED: cannot create $dest"; pause; exit 1; }

cp -f "$src" "$dest/Sankey.xlam" || { echo "FAILED: cannot copy to $dest (allow Terminal access to other apps' data in System Settings > Privacy & Security)"; pause; exit 1; }
xattr -dr com.apple.quarantine "$dest/Sankey.xlam" 2>/dev/null

echo "Installed: $dest/Sankey.xlam"
echo
echo "Start Excel: a Sankey tab appears next to Home."
echo "First start: if Excel asks about macros, choose Enable Macros."
echo "If the tab is missing, check Excel > Preferences > Security that macros are not fully disabled."
pause
