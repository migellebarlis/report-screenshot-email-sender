#!/bin/zsh
# Double-click in Finder (or run ./run.command) to send the report.
# Optional: ./run.command --date 2026-10-08 to send the report for another day.

ENV_NAME="report-screenshot-email-sender"
cd "$(dirname "$0")" || exit 1

# Find the conda env's Python without relying on the shell profile
PYTHON=""
for base in "$CONDA_PREFIX/.." /opt/miniconda3 /opt/anaconda3 "$HOME/miniconda3" "$HOME/anaconda3" /opt/homebrew/Caskroom/miniconda/base; do
    if [ -x "$base/envs/$ENV_NAME/bin/python" ]; then
        PYTHON="$base/envs/$ENV_NAME/bin/python"
        break
    fi
done

if [ -z "$PYTHON" ]; then
    echo "Could not find the '$ENV_NAME' conda env. See README.md to create it."
    read -k1 "?Press any key to close..."
    exit 1
fi

args=("$@")
# When double-clicked (no arguments, in a terminal), ask for an optional date
if [ ${#args[@]} -eq 0 ] && [ -t 0 ]; then
    read "report_date?Report date (YYYY-MM-DD), or press Enter for today: "
    [ -n "$report_date" ] && args=(--date "$report_date")
fi

"$PYTHON" main.py "${args[@]}"
exit_code=$?

echo
if [ $exit_code -eq 0 ]; then echo "Finished."; else echo "Exited with code $exit_code."; fi
# Keep the window open when double-clicked so the log can be read
[ -t 0 ] && read -k1 "?Press any key to close..."
exit $exit_code
