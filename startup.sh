#!/bin/sh
cd /workspace || exit 0

if curl -sf -o /dev/null --max-time 2 http://127.0.0.1:8080/; then
  exit 0
fi

python3 /workspace/server_download.py >/tmp/download-server.log 2>&1 &
sleep 0.5
exit 0
