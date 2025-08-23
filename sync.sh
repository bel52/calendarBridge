#!/bin/bash
# Use shell parameter expansion to get the script's directory.
# This is more robust than using the external 'dirname' command.
cd "${0%/*}" || exit

# Activate virtual environment and run the sync script.
./.venv/bin/python3 calendar_sync.py
