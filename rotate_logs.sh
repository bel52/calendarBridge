#!/bin/bash
cd ~/calendarBridge/logs || exit
timestamp=$(date +"%Y-%m")
for f in stdout.log stderr.log; do
  [ -f "$f" ] && mv "$f" "${f%.log}_$timestamp.log"
done
