#!/usr/bin/env bash
# 会意助手 · Knowing Assistant — 一键启动（macOS / Linux）
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
WRK="${TMPDIR:-/tmp}/knowing-assistant-run-$$"
mkdir -p "$WRK"

clear 2>/dev/null || true

cat <<'BANNER'

    ██╗  ██╗ ███╗   ██╗  ██████╗ ██╗    ██╗██╗███╗   ██╗ ██████╗
    ██║ ██╔╝ ████╗  ██║ ██╔════╝ ██║    ██║██║████╗  ██║██╔════╝
    █████╔╝  ██╔██╗ ██║ ██║  ███╗██║ █╗ ██║██║██╔██╗ ██║██║  ███╗
    ██╔═██╗  ██║╚██╗██║ ██║   ██║██║███╗██║██║██║╚██╗██║██║   ██║
    ██║  ██╗ ██║ ╚████║ ╚██████╔╝╚███╔███╔╝██║██║ ╚████║╚██████╔╝
    ╚═╝  ╚═╝ ╚═╝  ╚═══╝  ╚═════╝  ╚══╝╚══╝ ╚═╝╚═╝  ╚═══╝ ╚═════╝

              Knowing Assistant  ·  会意助手
              Your cognitive copilot for research & work

  ------------------------------------------------------------------

BANNER

if ! command -v node >/dev/null 2>&1; then
  echo "[×] 未检测到 Node.js，请先安装：https://nodejs.org/"
  echo "    建议版本：Node.js >= 22（与 backend/package.json engines 一致）"
  exit 1
fi

echo "[√] 检测到 Node.js $(node -v)"
echo "[*] 正在并发启动后端与前端（各开一个终端窗口）…"
echo

BACKEND_SCRIPT="$WRK/backend.sh"
FRONTEND_SCRIPT="$WRK/frontend.sh"

cat >"$BACKEND_SCRIPT" <<EOF
#!/usr/bin/env bash
set -e
cd "$ROOT/backend"
if [ ! -d node_modules ]; then
  echo "[backend] 首次安装依赖…"
  npm install
fi
echo "[backend] 启动开发服务 (npm run dev)…"
npm run dev
EOF
chmod +x "$BACKEND_SCRIPT"

cat >"$FRONTEND_SCRIPT" <<EOF
#!/usr/bin/env bash
set -e
cd "$ROOT/frontend"
if [ -f package-lock.json ]; then
  echo "[frontend] npm ci…"
  npm ci
else
  echo "[frontend] npm install…"
  npm install
fi
echo "[frontend] 启动开发服务 (npm run dev)…"
npm run dev
EOF
chmod +x "$FRONTEND_SCRIPT"

run_in_terminal() {
  local script_path="$1"
  local title="$2"
  if [[ "$(uname -s)" == "Darwin" ]]; then
    osascript \
      -e "tell application \"Terminal\"" \
      -e "  activate" \
      -e "  do script \"cd \\\"$ROOT\\\" && bash \\\"$script_path\\\"\"" \
      -e "end tell" >/dev/null
  else
    if command -v gnome-terminal >/dev/null 2>&1; then
      gnome-terminal --title="$title" -- bash -lc "bash '$script_path'; exec bash"
    elif command -v x-terminal-emulator >/dev/null 2>&1; then
      x-terminal-emulator -e bash -lc "bash '$script_path'; exec bash" &
    elif command -v xterm >/dev/null 2>&1; then
      xterm -title "$title" -e bash -lc "bash '$script_path'; exec bash" &
    elif command -v konsole >/dev/null 2>&1; then
      konsole -e bash -lc "bash '$script_path'; exec bash" &
    else
      echo "[×] 未找到可用终端（尝试安装 gnome-terminal / xterm，或使用 macOS Terminal）。"
      echo "    可手动执行: bash \"$script_path\""
      exit 1
    fi
  fi
}

run_in_terminal "$BACKEND_SCRIPT" "Knowing Assistant - Backend"
run_in_terminal "$FRONTEND_SCRIPT" "Knowing Assistant - Frontend"

echo "[√] 已请求打开两个终端窗口"

VITE_PORT="${VITE_PORT:-5173}"

wait_for_vite_port() {
  local port="${1:?}"
  local max_attempts="${2:-600}"
  local i=0
  echo "[*] 等待 Vite 就绪（检测 127.0.0.1:${port}，与 Frontend 里 「VITE … ready」/「Local:」一致后再打开浏览器）"
  echo "    首次运行 npm install 可能较慢，最长约等待 $(( max_attempts / 2 )) 秒…"
  while (( i < max_attempts )); do
    if (echo >/dev/tcp/127.0.0.1/"$port") 2>/dev/null; then
      return 0
    fi
    sleep 0.5
    i=$((i + 1))
  done
  return 1
}

URL="http://localhost:${VITE_PORT}"
if wait_for_vite_port "$VITE_PORT" 600; then
  echo "[√] 已检测到前端端口就绪，正在打开浏览器…"
  if [[ "$(uname -s)" == "Darwin" ]]; then
    open "$URL" || true
  else
    xdg-open "$URL" 2>/dev/null || sensible-browser "$URL" 2>/dev/null || true
  fi
else
  echo
  echo "【提示】在约定时间内仍未检测到端口 ${VITE_PORT}。"
  echo "        请查看 Frontend 终端：若仍在安装依赖，请继续等待直至出现："
  echo "          VITE v…  ready in … ms"
  echo "          Local:   http://localhost:${VITE_PORT}/"
  echo "        然后再手动在浏览器打开上述地址。"
  echo
fi

echo
echo "✅ 会意助手启动流程已结束。如需查看后端日志，请查看 Backend 终端窗口。"
echo
