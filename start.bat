@echo off
setlocal EnableExtensions EnableDelayedExpansion
cd /d "%~dp0"

cls
echo.
echo              HuiYi Assistant
echo              https://github.com/Xiaoxi753-code/HuiYiAssistant
echo.
echo  ------------------------------------------------------------------
echo.

rem ========== Step 1: check Node.js ==========
where node >nul 2>&1
if errorlevel 1 (
    echo  [x] Step 1/4 failed: Node.js not found
    echo.
    echo      Please do the following:
    echo        1^) Install Node.js LTS from https://nodejs.org/  suggested ^>= 22
    echo        2^) Check the "Add to PATH" option during install
    echo        3^) Close this window, open a new cmd, run:  node -v
    echo        4^) Then run start.bat again
    echo.
    pause
    exit /b 1
)
for /f "tokens=*" %%v in ('node -v 2^>nul') do set "NODE_VER=%%v"
echo  [OK] Step 1/4: Node.js detected  !NODE_VER!

rem ========== Step 2: check npm ==========
where npm >nul 2>&1
if errorlevel 1 (
    echo  [x] Step 2/4 failed: npm not found
    echo.
    echo      npm is included with Node.js. Try:
    echo        1^) Repair or reinstall Node.js
    echo        2^) Ensure the Node install folder is in PATH
    echo        3^) Open a new cmd and run:  npm -v
    echo.
    pause
    exit /b 1
)
for /f "tokens=*" %%v in ('npm -v 2^>nul') do set "NPM_VER=%%v"
echo  [OK] Step 2/4: npm detected  !NPM_VER!

rem ========== Step 3: check project structure ==========
set "ROOT=%~dp0"
if not exist "%ROOT%backend\package.json" (
    echo  [x] Step 3/4 failed: backend\package.json not found
    echo      Make sure start.bat is in the project root, same level as backend and frontend.
    echo.
    pause
    exit /b 1
)
if not exist "%ROOT%frontend\package.json" (
    echo  [x] Step 3/4 failed: frontend\package.json not found
    echo      Ensure the repo is fully cloned/unzipped and frontend exists.
    echo.
    pause
    exit /b 1
)
echo  [OK] Step 3/4: backend / frontend folders found

rem ========== Step 4: generate child scripts ==========
set "WRK=%TEMP%\huiyi-assistant-run"
if not exist "%WRK%" mkdir "%WRK%" 2>nul

set "BE=%WRK%\_backend.bat"
set "FE=%WRK%\_frontend.bat"

rem --- Backend script: write line by line to avoid redirection block issues ---
>"%BE%" echo @echo off
>>"%BE%" echo setlocal
>>"%BE%" echo title HuiYi Assistant - Backend
>>"%BE%" echo cd /d "%ROOT%backend"
>>"%BE%" echo if not exist node_modules ^(
>>"%BE%" echo   echo [backend] Installing dependencies for the first time...
>>"%BE%" echo   call npm install
>>"%BE%" echo   if errorlevel 1 ^(
>>"%BE%" echo     echo.
>>"%BE%" echo     echo [x] Backend npm install failed. Check network/proxy/disk/Node version.
>>"%BE%" echo     pause
>>"%BE%" echo     exit /b 1
>>"%BE%" echo   ^)
>>"%BE%" echo ^)
>>"%BE%" echo echo.
>>"%BE%" echo echo [backend] Starting npm run dev ...
>>"%BE%" echo call npm run dev
>>"%BE%" echo echo.
>>"%BE%" echo echo [backend] Process exited. Scroll up for errors if any.
>>"%BE%" echo pause

rem --- Frontend script: fallback from npm ci to npm install on failure ---
>"%FE%" echo @echo off
>>"%FE%" echo setlocal
>>"%FE%" echo title HuiYi Assistant - Frontend
>>"%FE%" echo cd /d "%ROOT%frontend"
>>"%FE%" echo if exist package-lock.json ^(
>>"%FE%" echo   echo [frontend] package-lock.json found, running npm ci ...
>>"%FE%" echo   call npm ci
>>"%FE%" echo   if errorlevel 1 ^(
>>"%FE%" echo     echo.
>>"%FE%" echo     echo [hint] npm ci failed, trying npm install ...
>>"%FE%" echo     call npm install
>>"%FE%" echo     if errorlevel 1 goto :fe_deps_fail
>>"%FE%" echo   ^)
>>"%FE%" echo ^) else ^(
>>"%FE%" echo   echo [frontend] package-lock.json not found, running npm install ...
>>"%FE%" echo   call npm install
>>"%FE%" echo   if errorlevel 1 goto :fe_deps_fail
>>"%FE%" echo ^)
>>"%FE%" echo goto :fe_run
>>"%FE%" echo :fe_deps_fail
>>"%FE%" echo echo.
>>"%FE%" echo echo [x] Frontend deps install failed. Check network/proxy or delete frontend\node_modules and retry.
>>"%FE%" echo pause
>>"%FE%" echo exit /b 1
>>"%FE%" echo :fe_run
>>"%FE%" echo echo.
>>"%FE%" echo echo [frontend] Starting npm run dev ...
>>"%FE%" echo call npm run dev
>>"%FE%" echo echo.
>>"%FE%" echo echo [frontend] Process exited. If the page does not load, check this window for errors.
>>"%FE%" echo pause

if not exist "%BE%" (
    echo  [x] Failed to write temp script: %BE%
    pause
    exit /b 1
)
if not exist "%FE%" (
    echo  [x] Failed to write temp script: %FE%
    pause
    exit /b 1
)

echo  [OK] Step 4/4: local launch scripts created
echo.
echo  [*] Opening 2 windows: Backend and Frontend. Please keep them open.
echo  [*] If you see only one window, the other may be grouped in the taskbar.
echo.

rem Parallel launch: two start commands, 1-second gap to avoid missing the second start
start "HuiYi Assistant - Backend" cmd /k call "%BE%"
ping -n 2 127.0.0.1 >nul
start "HuiYi Assistant - Frontend" cmd /k call "%FE%"

if errorlevel 1 (
    echo  [x] Failed to open the Frontend window (start command failed).
    echo      You can run manually:  %FE%
    echo.
    pause
    exit /b 1
)

echo  [OK] Backend and Frontend consoles requested (titles: HuiYi Assistant - Backend / Frontend)

set "VITE_PORT=5173"
echo.
echo  [*] Waiting for Vite to be ready (checking port !VITE_PORT!).
echo      First run may take a while while npm installs deps (up to ~5 minutes).
echo.

powershell -NoProfile -ExecutionPolicy Bypass -Command "$port=%VITE_PORT%; $max=600; $ok=$false; for ($i=0; $i -lt $max; $i++) { try { $t = New-Object System.Net.Sockets.TcpClient; $t.ReceiveTimeout=800; $t.SendTimeout=800; $t.Connect('127.0.0.1', $port); if ($t.Connected) { $t.Close(); $ok=$true; break } } catch { }; Start-Sleep -Milliseconds 500 }; if ($ok) { exit 0 } else { exit 1 }"

if errorlevel 1 (
    echo.
    echo  [hint] Port !VITE_PORT! not detected within the wait time.
    echo        Check the Frontend window. If deps are still installing, wait for:
    echo          VITE v...  ready in ... ms
    echo          Local:   http://localhost:!VITE_PORT!/
    echo        Then open that URL manually.
    echo.
) else (
    echo  [OK] Frontend port is ready. Opening browser...
    start "" "http://localhost:!VITE_PORT!" 2>nul
    if errorlevel 1 (
        echo  [hint] Could not open a browser. Please visit: http://localhost:!VITE_PORT!
        echo.
    )
)

echo.
echo  [done] Launch sequence finished. For backend logs, see the Backend window.
echo     If frontend did not start, check the Frontend window for errors.
echo.
pause
endlocal
