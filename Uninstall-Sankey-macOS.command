#!/bin/bash
# Removes the Sankey add-in from Excel on macOS. Other add-ins are not touched.
# Double-click this file, or run:  bash Uninstall-Sankey-macOS.command
pause() { echo; read -r -p "Press Enter to close" _; }

if pgrep -x "Microsoft Excel" >/dev/null 2>&1; then
  echo "FAILED: Excel is running. Save your work, quit Excel (Cmd+Q) and run this again."
  pause; exit 1
fi

read -r -p "Remove the Sankey add-in from Excel? (y/N) " a
case "$a" in y|Y|yes|YES) ;; *) echo "Cancelled - nothing was changed."; pause; exit 0;; esac

container="$HOME/Library/Group Containers/UBF8T346G9.Office"
removed=0
for uc in "$container/User Content.localized" "$container/User Content"; do
  for f in "$uc/Startup.localized/Excel/Sankey.xlam" "$uc/Startup/Excel/Sankey.xlam" \
           "$uc/Add-Ins.localized/Sankey.xlam" "$uc/Add-Ins/Sankey.xlam"; do
    if [ -e "$f" ]; then
      rm -f "$f" && echo "Deleted $f" && removed=$((removed+1))
    fi
  done
done
[ "$removed" -eq 0 ] && echo "No Sankey add-in file found (nothing to delete)."
echo
echo "Done. Charts already in your workbooks stay as ordinary shapes. Restart Excel to confirm the Sankey tab is gone."
pause
