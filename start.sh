#!/bin/bash
echo "[start] Launching Discord bot..."
python3 bot.py &
BOT_PID=$!

echo "[start] Launching Flask render server on port $PORT..."
gunicorn -w 1 -b 0.0.0.0:$PORT --timeout 60 render_server:app &
FLASK_PID=$!

wait -n $BOT_PID $FLASK_PID
echo "[start] Process exited, shutting down..."
kill $BOT_PID $FLASK_PID 2>/dev/null
exit 1
