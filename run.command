#!/bin/zsh
# Double-click in Finder (or run ./run.command) to send the report.

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

"$PYTHON" main.py
status=$?

echo
if [ $status -eq 0 ]; then echo "Finished."; else echo "Exited with code $status."; fi
# Keep the window open when double-clicked so the log can be read
[ -t 0 ] && read -k1 "?Press any key to close..."
exit $status
