#!/usr/bin/env bash
set -euo pipefail

SUPPORT_SYSTEM_DLL_PATH="${SUPPORT_SYSTEM_DLL_PATH:-/app/ITSupportSystem/publish/ITSupportSystem.dll}"

if [ ! -f "$SUPPORT_SYSTEM_DLL_PATH" ]; then
  echo "ITSupportSystem.dll blev ikke fundet: $SUPPORT_SYSTEM_DLL_PATH"
  exit 1
fi

echo "Starter ITSupportSystem monitor..."
dotnet "$SUPPORT_SYSTEM_DLL_PATH" run &
MONITOR_PID=$!

cleanup() {
  if kill -0 "$MONITOR_PID" >/dev/null 2>&1; then
    kill "$MONITOR_PID" >/dev/null 2>&1 || true
  fi
}

trap cleanup EXIT INT TERM

echo "Starter Website server..."
cd /app/Website
node server.js
