@echo off
setlocal EnableExtensions EnableDelayedExpansion
chcp 65001 >nul
cd /d "%~dp0"

cls
echo.
echo    ██╗  ██╗ ███╗   ██╗  ██████╗ ██╗    ██╗██╗███╗   ██╗ ██████╗
echo    ██║ ██╔╝ ████╗  ██║ ██╔════╝ ██║    ██║██║████╗  ██║██╔════╝
echo    █████╔╝  ██╔██╗ ██║ ██║  ███╗██║ █╗ ██║██║██╔██╗ ██║██║  ███╗
echo    ██╔═██╗  ██║╚██╗██║ ██║   ██║██║███╗██║██║██║╚██╗██║██║   ██║
echo    ██║  ██╗ ██║ ╚████║ ╚██████╔╝╚███╔███╔╝██║██║ ╚████║╚██████╔╝
echo    ╚═╝  ╚═╝ ╚═╝  ╚═══╝  ╚═════╝  ╚══╝╚══╝ ╚═╝╚═╝  ╚═══╝ ╚═════╝
echo.
echo              Knowing Assistant  ·  会意助手
echo              Your cognitive copilot for research ^& work
echo.
echo  ------------------------------------------------------------------
echo.

rem ========== 步骤 1：检查 Node.js ==========
where node >nul 2>&1
if errorlevel 1 (
    echo  [×] 步骤 1/4 失败：未找到 Node.js
    echo.
    echo      请按下面顺序操作：
    echo        1^) 打开 https://nodejs.org/ 下载并安装 LTS（建议 ^>= 22^）
    echo        2^) 安装时勾选 “Add to PATH”
    echo        3^) **关闭本窗口**，重新打开命令提示符，执行：  node -v
    echo        4^) 看到版本号后，再次双击 start.bat
    echo.
    pause
    exit /b 1
)
for /f "tokens=*" %%v in ('node -v 2^>nul') do set "NODE_VER=%%v"
echo  [√] 步骤 1/4：Node.js 已安装  !NODE_VER!

rem ========== 步骤 2：检查 npm（随 Node 自带，单独确认）==========
where npm >nul 2>&1
if errorlevel 1 (
    echo  [×] 步骤 2/4 失败：未找到 npm
    echo.
    echo      npm 通常随 Node 一起安装。请尝试：
    echo        1^) 修复安装 Node.js，或重装 Node.js
    echo        2^) 以管理员身份打开「环境变量」，确认 Node 安装目录在 PATH 中
    echo        3^) 新开 cmd 执行：  npm -v
    echo.
    pause
    exit /b 1
)
for /f "tokens=*" %%v in ('npm -v 2^>nul') do set "NPM_VER=%%v"
echo  [√] 步骤 2/4：npm 可用  !NPM_VER!

rem ========== 步骤 3：检查项目目录结构 ==========
set "ROOT=%~dp0"
if not exist "%ROOT%backend\package.json" (
    echo  [×] 步骤 3/4 失败：找不到 backend\package.json
    echo      请确认 start.bat 放在项目根目录（与 backend、frontend 文件夹同级）。
    echo.
    pause
    exit /b 1
)
if not exist "%ROOT%frontend\package.json" (
    echo  [×] 步骤 3/4 失败：找不到 frontend\package.json
    echo      请确认已完整克隆/解压仓库，且 frontend 目录存在。
    echo.
    pause
    exit /b 1
)
echo  [√] 步骤 3/4：backend / frontend 目录正常

rem ========== 步骤 4：生成子脚本（避免括号块写入损坏第二份脚本）==========
set "WRK=%TEMP%\knowing-assistant-run"
if not exist "%WRK%" mkdir "%WRK%" 2>nul

set "BE=%WRK%\_backend.bat"
set "FE=%WRK%\_frontend.bat"

rem --- 后端启动脚本（逐行写入，禁止用 ^> file ^( ... ^) 多行重定向）---
>"%BE%" echo @echo off
>>"%BE%" echo setlocal
>>"%BE%" echo chcp 65001 ^>nul
>>"%BE%" echo title Knowing Assistant - Backend
>>"%BE%" echo cd /d "%ROOT%backend"
>>"%BE%" echo if not exist node_modules ^(
>>"%BE%" echo   echo [backend] 首次安装依赖，请稍候…
>>"%BE%" echo   call npm install
>>"%BE%" echo   if errorlevel 1 ^(
>>"%BE%" echo     echo.
>>"%BE%" echo     echo [×] 后端 npm install 失败。请检查：网络、代理、磁盘空间、Node 版本。
>>"%BE%" echo     pause
>>"%BE%" echo     exit /b 1
>>"%BE%" echo   ^)
>>"%BE%" echo ^)
>>"%BE%" echo echo.
>>"%BE%" echo echo [backend] 启动 npm run dev …
>>"%BE%" echo call npm run dev
>>"%BE%" echo echo.
>>"%BE%" echo echo [backend] 进程已结束。若曾报错请向上滚动查看。
>>"%BE%" echo pause

rem --- 前端启动脚本（npm ci 失败则自动改 npm install，全程用 errorlevel，避免延迟变量）---
>"%FE%" echo @echo off
>>"%FE%" echo setlocal
>>"%FE%" echo chcp 65001 ^>nul
>>"%FE%" echo title Knowing Assistant - Frontend
>>"%FE%" echo cd /d "%ROOT%frontend"
>>"%FE%" echo if exist package-lock.json ^(
>>"%FE%" echo   echo [frontend] 检测到 package-lock.json，执行 npm ci …
>>"%FE%" echo   call npm ci
>>"%FE%" echo   if errorlevel 1 ^(
>>"%FE%" echo     echo.
>>"%FE%" echo     echo 【提示】npm ci 未成功（锁与依赖不一致较常见），改试 npm install …
>>"%FE%" echo     call npm install
>>"%FE%" echo     if errorlevel 1 goto :fe_deps_fail
>>"%FE%" echo   ^)
>>"%FE%" echo ^) else ^(
>>"%FE%" echo   echo [frontend] 未找到 package-lock.json，执行 npm install …
>>"%FE%" echo   call npm install
>>"%FE%" echo   if errorlevel 1 goto :fe_deps_fail
>>"%FE%" echo ^)
>>"%FE%" echo goto :fe_run
>>"%FE%" echo :fe_deps_fail
>>"%FE%" echo echo.
>>"%FE%" echo echo [×] 前端依赖安装失败。请检查网络/代理，或删除 frontend\node_modules 后重试。
>>"%FE%" echo pause
>>"%FE%" echo exit /b 1
>>"%FE%" echo :fe_run
>>"%FE%" echo echo.
>>"%FE%" echo echo [frontend] 启动 npm run dev …
>>"%FE%" echo call npm run dev
>>"%FE%" echo echo.
>>"%FE%" echo echo [frontend] 进程已结束。若页面打不开，请确认本窗口无报错。
>>"%FE%" echo pause

if not exist "%BE%" (
    echo  [×] 无法写入临时脚本：%BE%
    pause
    exit /b 1
)
if not exist "%FE%" (
    echo  [×] 无法写入临时脚本：%FE%
    pause
    exit /b 1
)

echo  [√] 步骤 4/4：已生成本地启动脚本
echo.
echo  [*] 即将打开 2 个窗口：【后端】与【前端】。请勿关闭它们。
echo  [*] 若只看到 1 个窗口，请阅读该窗口内提示；另一个可能被任务栏合并，请展开任务栏分组。
echo.

rem 并发启动：两次 start 互不等待；中间间隔 1 秒，避免个别环境吞掉第二次 start
start "Knowing Assistant - Backend" cmd /k call "%BE%"
ping -n 2 127.0.0.1 >nul
start "Knowing Assistant - Frontend" cmd /k call "%FE%"

if errorlevel 1 (
    echo  [×] 启动前端窗口失败（start 命令异常）。
    echo      请手动双击运行：  %FE%
    echo.
    pause
    exit /b 1
)

echo  [√] 已请求打开后端与前端控制台（标题分别为 Knowing Assistant - Backend / Frontend）

set "VITE_PORT=5173"
echo.
echo  [*] 正在等待 Vite 开发服务器就绪（检测本机端口 !VITE_PORT!，与 Frontend 窗口里 “VITE … ready” 一致后再打开浏览器）
echo      若首次运行，npm install 可能较慢，请稍候（最长约等待 5 分钟）…
echo.

powershell -NoProfile -ExecutionPolicy Bypass -Command "$port=%VITE_PORT%; $max=600; $ok=$false; for ($i=0; $i -lt $max; $i++) { try { $t = New-Object System.Net.Sockets.TcpClient; $t.ReceiveTimeout=800; $t.SendTimeout=800; $t.Connect('127.0.0.1', $port); if ($t.Connected) { $t.Close(); $ok=$true; break } } catch { }; Start-Sleep -Milliseconds 500 }; if ($ok) { exit 0 } else { exit 1 }"

if errorlevel 1 (
    echo.
    echo  【提示】在约定时间内仍未检测到端口 !VITE_PORT! 监听。
    echo        请查看 「Knowing Assistant - Frontend」窗口：若仍在安装依赖，请继续等待直至出现：
    echo          VITE v…  ready in … ms
    echo          Local:   http://localhost:!VITE_PORT!/
    echo        然后再手动在浏览器打开上述地址。
    echo.
) else (
    echo  [√] 已检测到前端端口就绪，正在打开浏览器…
    start "" "http://localhost:!VITE_PORT!" 2>nul
    if errorlevel 1 (
        echo  【提示】系统未能自动唤起浏览器，请手动访问： http://localhost:!VITE_PORT!
        echo.
    )
)

echo.
echo  ✅ 会意助手启动流程已执行！如需查看后端日志，请查看标题为 Backend 的控制台窗口。
echo     若前端未启动：请看 Frontend 窗口中的报错，并按其中提示操作。
echo.
pause
endlocal
