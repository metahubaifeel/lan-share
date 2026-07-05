#!/bin/bash
# LAN Share Hub — one-click start (local LAN mode)
set -e

PORT="${HUB_PORT:-8888}"
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

fuser -k "${PORT}/tcp" 2>/dev/null || true
sleep 0.5

cd "$SCRIPT_DIR"

python3 server.py &
SERVER_PID=$!
sleep 1.5

LAN_IP=$(python3 -c "
import subprocess
for ip in subprocess.check_output(['hostname','-I'], text=True).split():
    if ip.startswith('192.168.') or ip.startswith('10.'):
        print(ip)
        break
" 2>/dev/null || echo "127.0.0.1")

xdg-open "http://localhost:${PORT}/" 2>/dev/null || true

echo ""
echo "  ✓ LAN Share Hub started"
echo "  PC:     http://localhost:${PORT}/"
echo "  Phone:  http://${LAN_IP}:${PORT}/  (same WiFi required)"
echo "  PID:    ${SERVER_PID}"
echo ""
echo "  Stop:   kill ${SERVER_PID}"
echo "  Tip:    IP changes when WiFi changes — copy URL from the top bar"
echo ""
