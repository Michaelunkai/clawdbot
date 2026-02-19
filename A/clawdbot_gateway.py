"""
Clawdbot Gateway - System Tray Application
Runs the gateway in background without terminal window
"""

import os
import sys
import subprocess
import threading
from pathlib import Path

# Handle PyInstaller frozen exe
if getattr(sys, 'frozen', False):
    APP_DIR = Path(sys._MEIPASS)
    EXE_DIR = Path(sys.executable).parent
else:
    APP_DIR = Path(__file__).parent
    EXE_DIR = APP_DIR

# Try importing pystray, with fallback message
try:
    import pystray
    from PIL import Image, ImageDraw
except ImportError:
    print("Missing dependencies. Run: pip install pystray pillow")
    sys.exit(1)


class ClawdbotGateway:
    def __init__(self):
        self.process = None
        self.icon = None
        self.running = False
        
    def create_icon_image(self, color='#4CAF50'):
        """Create a simple icon - green circle when running"""
        size = 64
        image = Image.new('RGBA', (size, size), (0, 0, 0, 0))
        draw = ImageDraw.Draw(image)
        # Draw filled circle
        margin = 4
        draw.ellipse([margin, margin, size-margin, size-margin], fill=color)
        # Draw "C" letter
        draw.text((size//2 - 8, size//2 - 12), "C", fill='white')
        return image
    
    def get_env(self):
        """Get environment variables for gateway"""
        env = os.environ.copy()
        env['PATH'] = (
            r'C:\Users\User\.local\bin;'
            r'C:\Users\User\.bun\bin;'
            r'F:\DevKit\dotnet;'
            r'C:\WINDOWS\system32;'
            r'C:\WINDOWS;'
            r'C:\WINDOWS\System32\Wbem;'
            r'F:\DevKit\tools\7zip;'
            r'C:\WINDOWS\System32\WindowsPowerShell\v1.0\;'
            r'C:\WINDOWS\System32\OpenSSH\;'
            r'C:\Program Files\PowerShell\7\;'
            r'C:\Program Files\nodejs\;'
            r'C:\Program Files\NVIDIA Corporation\NVIDIA App\NvDLISR;'
            r'C:\Program Files\dotnet\;'
            r'C:\ProgramData\chocolatey\bin;'
            r'C:\Program Files\Git\cmd;'
            r'C:\Program Files\Shield\;'
            r'C:\Program Files (x86)\NVIDIA Corporation\PhysX\Common;'
            r'C:\Program Files\GitHub CLI\;'
            r'C:\Program Files\Docker\Docker\resources\bin;'
            r'C:\Users\User\AppData\Local\Programs\Python\Python312\Scripts\;'
            r'C:\Users\User\AppData\Local\Programs\Python\Python312\;'
            r'C:\Users\User\AppData\Local\Programs\Python\Launcher\;'
            r'C:\Users\User\AppData\Local\Microsoft\WindowsApps;'
            r'C:\Users\User\AppData\Roaming\npm;'
            r'C:\Users\User\AppData\Local\Microsoft\WinGet\Packages\Game1024.OpenSpeedy_Microsoft.Winget.Source_8wekyb3d8bbwe;'
            r'C:\Users\User\AppData\Local\Microsoft\WinGet\Links;'
            r'C:\Users\User\AppData\Local\GitHubDesktop\bin;'
            r'C:\Users\User\AppData\Local\Programs\Antigravity\bin;'
            r'C:\Users\User\.dotnet\tools;'
            r'C:\winlibs\mingw64\bin;'
            r'F:\backup\Nodejs\global;'
            r'C:\Users\User\bin'
        )
        env['CLAWDBOT_GATEWAY_PORT'] = '18789'
        env['CLAWDBOT_GATEWAY_TOKEN'] = '29da324b6c9cd281c0179a2f9995946d78a9e719c0f8d27c'
        env['CLAWDBOT_SYSTEMD_UNIT'] = 'clawdbot-gateway.service'
        env['CLAWDBOT_SERVICE_MARKER'] = 'clawdbot'
        env['CLAWDBOT_SERVICE_KIND'] = 'gateway'
        env['CLAWDBOT_SERVICE_VERSION'] = '2026.1.24-3'
        return env
    
    def start_gateway(self):
        """Start the gateway process"""
        if self.process and self.process.poll() is None:
            return  # Already running
            
        cmd = [
            r'C:\Program Files\nodejs\node.exe',
            r'C:\Users\User\AppData\Roaming\npm\node_modules\clawdbot\dist\entry.js',
            'gateway',
            '--port', '18789'
        ]
        
        # Start without console window
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        startupinfo.wShowWindow = subprocess.SW_HIDE
        
        self.process = subprocess.Popen(
            cmd,
            env=self.get_env(),
            startupinfo=startupinfo,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            creationflags=subprocess.CREATE_NO_WINDOW
        )
        self.running = True
        self.update_icon()
        
    def stop_gateway(self):
        """Stop the gateway process"""
        if self.process:
            self.process.terminate()
            try:
                self.process.wait(timeout=5)
            except subprocess.TimeoutExpired:
                self.process.kill()
            self.process = None
        self.running = False
        self.update_icon()
        
    def restart_gateway(self, icon=None, item=None):
        """Restart the gateway"""
        self.stop_gateway()
        self.start_gateway()
        
    def update_icon(self):
        """Update icon color based on status"""
        if self.icon:
            color = '#4CAF50' if self.running else '#F44336'  # Green or Red
            self.icon.icon = self.create_icon_image(color)
            
    def on_quit(self, icon, item):
        """Quit the application"""
        self.stop_gateway()
        icon.stop()
        
    def create_menu(self):
        """Create system tray menu"""
        return pystray.Menu(
            pystray.MenuItem('Clawdbot Gateway', None, enabled=False),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem('Restart', self.restart_gateway),
            pystray.MenuItem('Quit', self.on_quit)
        )
    
    def run(self):
        """Main entry point"""
        # Start gateway
        self.start_gateway()
        
        # Create system tray icon
        self.icon = pystray.Icon(
            'clawdbot',
            self.create_icon_image('#4CAF50'),
            'Clawdbot Gateway',
            self.create_menu()
        )
        
        # Run the icon (blocking)
        self.icon.run()


if __name__ == '__main__':
    app = ClawdbotGateway()
    app.run()
