@echo off
rem Clawdbot Gateway (v2026.1.24-3)
set PATH=C:\Users\User\.local\bin;C:\Users\User\.bun\bin;F:\DevKit\dotnet;C:\WINDOWS\system32;C:\WINDOWS;C:\WINDOWS\System32\Wbem;F:\DevKit\tools\7zip;C:\WINDOWS\System32\WindowsPowerShell\v1.0\;C:\WINDOWS\System32\OpenSSH\;C:\Program Files\PowerShell\7\;C:\Program Files\nodejs\;C:\Program Files\NVIDIA Corporation\NVIDIA App\NvDLISR;C:\Program Files\dotnet\;C:\ProgramData\chocolatey\bin;C:\Program Files\Git\cmd;C:\Program Files\Shield\;C:\Program Files (x86)\NVIDIA Corporation\PhysX\Common;C:\Program Files\GitHub CLI\;C:\Program Files\Docker\Docker\resources\bin;C:\Users\User\AppData\Local\Programs\Python\Python312\Scripts\;C:\Users\User\AppData\Local\Programs\Python\Python312\;C:\Users\User\AppData\Local\Programs\Python\Launcher\;C:\Users\User\AppData\Local\Microsoft\WindowsApps;C:\Users\User\AppData\Roaming\npm;C:\Users\User\AppData\Local\Microsoft\WinGet\Packages\Game1024.OpenSpeedy_Microsoft.Winget.Source_8wekyb3d8bbwe;C:\Users\User\AppData\Local\Microsoft\WinGet\Links;C:\Users\User\AppData\Local\GitHubDesktop\bin;C:\Users\User\AppData\Local\Programs\Antigravity\bin;C:\Users\User\.dotnet\tools;C:\winlibs\mingw64\bin;F:\\backup\\Nodejs\\global;C:\Users\User\bin
set CLAWDBOT_GATEWAY_PORT=18789
set CLAWDBOT_GATEWAY_TOKEN=29da324b6c9cd281c0179a2f9995946d78a9e719c0f8d27c
set CLAWDBOT_SYSTEMD_UNIT=clawdbot-gateway.service
set CLAWDBOT_SERVICE_MARKER=clawdbot
set CLAWDBOT_SERVICE_KIND=gateway
set CLAWDBOT_SERVICE_VERSION=2026.1.24-3
"C:\Program Files\nodejs\node.exe" C:\Users\User\AppData\Roaming\npm\node_modules\clawdbot\dist\entry.js gateway --port 18789
