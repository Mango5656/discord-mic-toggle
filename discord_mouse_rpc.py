"""
Discord Mouse Controller - Web UI Version
使用 PyWebView 提供精美的 Glassmorphism 介面
"""

import webview
import threading
import asyncio
from pypresence import AioClient
from pynput import mouse, keyboard
import time
import json
import os
import requests
import struct
import pystray
from PIL import Image, ImageDraw
import winreg
import sys

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(os.path.abspath(__file__))

    return os.path.join(base_path, relative_path)

CONFIG_FILE = "config.json"
HTML_FILE = resource_path(os.path.join("web", "index.html"))


class DiscordAPI:
    """Bridge between Web UI and Discord RPC"""
    
    def __init__(self):
        self.window = None
        self.rpc_client = None
        self.loop = None
        self.running = False
        self.current_voice_settings = {'deaf': False, 'mute': False}
        self.binding_target = None
        self.binding_pending = False  # Block action triggers during binding
        self.tray_icon = None
        
        # Config
        self.config = self.load_config()
        self.saved_access_token = self.config.get('access_token')
        self.saved_refresh_token = self.config.get('refresh_token')
        
        # Start input listeners
        self.mouse_listener = mouse.Listener(on_click=self.on_click)
        self.mouse_listener.start()
        
        self.keyboard_listener = keyboard.Listener(on_press=self.on_key_press)
        self.keyboard_listener.start()
    
    def set_window(self, window):
        """Set the webview window reference"""
        self.window = window
        # Load config into UI
        window.evaluate_js(f"loadConfig({json.dumps(self.config)})")
        
        # Auto-connect if credentials exist
        if self.config.get('client_id') and self.config.get('client_secret'):
            threading.Thread(target=self._delayed_connect, daemon=True).start()
    
    def _delayed_connect(self):
        time.sleep(1)
        self.connect(self.config.get('client_id'), self.config.get('client_secret'))
    
    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r') as f:
                    return json.load(f)
            except:
                pass
        return {}
    
    def save_config(self, data=None):
        """Save config from UI"""
        if data:
            self.config.update(data)
        
        # Always include tokens
        self.config['access_token'] = self.saved_access_token
        self.config['refresh_token'] = self.saved_refresh_token
        
        try:
            with open(CONFIG_FILE, 'w') as f:
                json.dump(self.config, f)
        except:
            pass

    def test_api(self):
        """Test if API is working"""
        print("[API] test_api was called successfully!")
        return "Connected"

    def close_window(self):
        """Close window - trigger the closing event"""
        self.running = False
        
        # Stop listeners immediately
        try:
            if self.mouse_listener:
                self.mouse_listener.stop()
        except:
            pass
        try:
            if self.keyboard_listener:
                self.keyboard_listener.stop()
        except:
            pass
        
        # Close asyncio loop
        if self.loop and self.loop.is_running():
            self.loop.call_soon_threadsafe(self.loop.stop)
        
        if self.window:
            # Use threading to avoid UI freeze
            def do_close():
                time.sleep(0.05)
                if self.window:
                    self.window.destroy()
            threading.Thread(target=do_close, daemon=True).start()
    
    def minimize_window(self):
        """Minimize window"""
        if self.window:
            self.window.minimize()
    
    def toggle_fullscreen(self):
        """Toggle fullscreen mode"""
        if self.window:
            self.window.toggle_fullscreen()

    def set_startup(self, enabled):
        """Set or unset app startup in Windows Registry"""
        key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
        app_name = "DiscordMouseController"
        exe_path = sys.executable if getattr(sys, 'frozen', False) else f'"{sys.executable}" "{os.path.abspath(sys.argv[0])}"'

        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_ALL_ACCESS)
            if enabled:
                winreg.SetValueEx(key, app_name, 0, winreg.REG_SZ, exe_path)
            else:
                try:
                    winreg.DeleteValue(key, app_name)
                except FileNotFoundError:
                    pass
            winreg.CloseKey(key)
            return True
        except Exception as e:
            print(f"Error setting startup: {e}")
            return False

    def get_startup(self):
        """Check if app is set to run at startup"""
        key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
        app_name = "DiscordMouseController"
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_READ)
            winreg.QueryValueEx(key, app_name)
            winreg.CloseKey(key)
            return True
        except FileNotFoundError:
            return False
        except Exception as e:
            print(f"Error checking startup: {e}")
            return False

    def toggle_mute(self):
        """Directly toggle mute status"""
        print("[API] Toggle Mute requested")
        self.trigger_action('mute')
        return True

    def toggle_deafen(self):
        """Directly toggle deafen status"""
        print("[API] Toggle Deafen requested")
        self.trigger_action('deafen')
        return True
    
    def start_drag(self):
        """Start window drag - for custom title bar"""
        if self.window:
            self.window.move(0, 0)  # This enables native drag
    
    def set_bind_target(self, target_type):
        """Set the binding target - 0 for deafen, 1 for mute"""
        print(f"[API] set_bind_target called with: {target_type}")
        
        # Immediately block action triggers
        self.binding_pending = True
        
        if target_type == 0:
            self.start_binding('deafen')
        else:
            self.start_binding('mute')
        return True

    def start_binding(self, target):
        """Start keyboard/mouse binding for a target"""
        # Force string and strip
        target = str(target).strip()
        
        print(f"[-DEBUG-] start_binding RAW target: '{target}' (len={len(target)})")
        
        if target == 'deafen':
            print("[-DEBUG-] Target MATCHES 'deafen'")
        elif target == 'mute':
            print("[-DEBUG-] Target MATCHES 'mute'")
        else:
            print(f"[-DEBUG-] Target '{target}' DOES NOT MATCH expected values!")

        self.binding_target = target
        print(f"[-DEBUG-] binding_target NOW SET TO: '{self.binding_target}'")
        
        # Explicit feedback to UI
        target_name = "拒聽" if target == 'deafen' else "靜音"
        self.update_status(f"Python 已就緒: 請按下 {target_name} 鍵...")
        
        print(f"[-DEBUG-] After update_status, binding_target is: '{self.binding_target}'")
    
    def connect(self, client_id, client_secret):
        """Connect to Discord RPC"""
        if not client_id or not client_secret:
            return
        
        self.config['client_id'] = client_id
        self.config['client_secret'] = client_secret
        self.save_config()
        
        self.running = True
        threading.Thread(target=self._run_rpc, args=(client_id, client_secret), daemon=True).start()
    
    def disconnect(self):
        """Disconnect from Discord RPC"""
        self.running = False
        self.update_connection_status(False)
        self.update_status("已斷開連接")
    
    def _run_rpc(self, client_id, client_secret):
        """Run the async RPC connection in a thread"""
        self.loop = asyncio.new_event_loop()
        asyncio.set_event_loop(self.loop)
        
        while self.running:
            try:
                self.loop.run_until_complete(self._async_main(client_id, client_secret))
                if self.running:
                    time.sleep(2)
            except Exception as e:
                print(f"RPC Error: {e}")
                if self.running:
                    time.sleep(2)
                else:
                    break
    
    async def _async_main(self, client_id, client_secret):
        """Main async RPC logic"""
        try:
            self.rpc_client = AioClient(client_id)
            
            # Connect with retries
            self.update_status("正在連接 Discord...")
            connected = False
            for i in range(3):
                try:
                    await self.rpc_client.start()
                    connected = True
                    break
                except Exception as e:
                    print(f"Connection attempt {i+1} failed: {e}")
                    await asyncio.sleep(1)
            
            if not connected:
                raise Exception("無法連接到 Discord，請確認 Discord 已啟動")
            
            # Auth
            access_token = self.saved_access_token
            refresh_token = self.saved_refresh_token
            
            if not access_token:
                self.update_status("請在 Discord 視窗中點擊授權...")
                auth_resp = await self.rpc_client.authorize(client_id, scopes=['rpc'])
                code = auth_resp['data']['code']
                
                self.update_status("交換 Token 中...")
                token_resp = requests.post(
                    'https://discord.com/api/oauth2/token',
                    data={
                        'client_id': client_id,
                        'client_secret': client_secret,
                        'grant_type': 'authorization_code',
                        'code': code,
                        'redirect_uri': 'http://127.0.0.1'
                    }
                )
                
                if token_resp.status_code != 200:
                    raise Exception(f"Token 交換失敗: {token_resp.text}")
                
                token_data = token_resp.json()
                access_token = token_data['access_token']
                refresh_token = token_data.get('refresh_token')
                
                self.saved_access_token = access_token
                self.saved_refresh_token = refresh_token
                self.save_config()
            
            # Authenticate
            self.update_status("驗證中...")
            try:
                await self.rpc_client.authenticate(access_token)
            except Exception as auth_error:
                print(f"Auth failed: {auth_error}")
                
                # Try refresh
                if refresh_token:
                    self.update_status("Token 已過期，正在刷新...")
                    refresh_resp = requests.post(
                        'https://discord.com/api/oauth2/token',
                        data={
                            'client_id': client_id,
                            'client_secret': client_secret,
                            'grant_type': 'refresh_token',
                            'refresh_token': refresh_token
                        }
                    )
                    
                    if refresh_resp.status_code == 200:
                        token_data = refresh_resp.json()
                        access_token = token_data['access_token']
                        refresh_token = token_data.get('refresh_token', refresh_token)
                        
                        self.saved_access_token = access_token
                        self.saved_refresh_token = refresh_token
                        self.save_config()
                        
                        await self.rpc_client.authenticate(access_token)
                    else:
                        # Need full re-auth
                        self.saved_access_token = None
                        self.saved_refresh_token = None
                        self.save_config()
                        raise Exception("驗證失敗，請重新連接")
                else:
                    raise auth_error
            
            # Subscribe to voice updates
            await self.rpc_client.subscribe('VOICE_SETTINGS_UPDATE')
            
            self.update_status("已連接")
            self.update_connection_status(True)
            
            # Read loop
            await self._read_loop()
            
        except Exception as e:
            print(f"Error: {e}")
            self.running = False
            self.update_status(f"錯誤: {str(e)}")
            self.update_connection_status(False)
        finally:
            if self.rpc_client and hasattr(self.rpc_client, 'sock_writer'):
                try:
                    self.rpc_client.sock_writer.close()
                except:
                    pass
    
    async def _read_loop(self):
        """Read discord events"""
        while self.running:
            try:
                data = await self.rpc_client.read_output()
                
                if data.get('cmd') == 'DISPATCH' and data.get('evt') == 'VOICE_SETTINGS_UPDATE':
                    payload = data.get('data', {})
                    self.current_voice_settings['deaf'] = payload.get('deaf', False)
                    self.current_voice_settings['mute'] = payload.get('mute', False)
                    self.update_voice_status()
                    
            except Exception as e:
                if "timed out" in str(e) or "No response" in str(e):
                    continue
                print(f"Read error: {e}")
                break
    
    async def _toggle_deaf(self):
        """Toggle deafen status"""
        try:
            new_deaf = not self.current_voice_settings['deaf']
            payload = {
                'cmd': 'SET_VOICE_SETTINGS',
                'args': {'deaf': new_deaf},
                'nonce': str(time.time())
            }
            await self._send_payload(1, payload)
        except Exception as e:
            print(f"Toggle deaf error: {e}")
    
    async def _toggle_mute(self):
        """Toggle mute status"""
        try:
            new_mute = not self.current_voice_settings['mute']
            payload = {
                'cmd': 'SET_VOICE_SETTINGS',
                'args': {'mute': new_mute},
                'nonce': str(time.time())
            }
            await self._send_payload(1, payload)
        except Exception as e:
            print(f"Toggle mute error: {e}")
    
    async def _send_payload(self, op, payload):
        """Send raw payload to Discord"""
        try:
            payload_json = json.dumps(payload)
            encoded = payload_json.encode('utf-8')
            header = struct.pack('<II', op, len(encoded))
            
            if hasattr(self.rpc_client, 'sock_writer'):
                self.rpc_client.sock_writer.write(header + encoded)
                await self.rpc_client.sock_writer.drain()
        except:
            pass
    
    # === Input Handlers ===
    
    def on_key_press(self, key):
        self._handle_input(str(key))
    
    def on_click(self, x, y, button, pressed):
        if pressed:
            self._handle_input(str(button))
    
    def _handle_input(self, input_id):
        print(f"[DEBUG] _handle_input called with: {input_id}")
        print(f"[DEBUG] current binding_target: {self.binding_target}")
        
        # Escape for safe JS injection
        safe_input = json.dumps(input_id)  # This adds quotes and escapes properly
        
        # Update UI
        if self.window:
            try:
                self.window.evaluate_js(f"updateLastInput({safe_input})")
            except:
                pass
        
        # Binding mode
        if self.binding_target:
            print(f"[DEBUG] In binding mode for target: {self.binding_target}")
            if input_id == "Button.left":
                print("[DEBUG] Ignoring left click")
                return
            
            target = self.binding_target
            self.binding_target = None
            
            if target == 'deafen':
                self.config['btn_deafen'] = input_id
                print(f"[DEBUG] Set btn_deafen to: {input_id}")
            elif target == 'mute':
                self.config['btn_mute'] = input_id
                print(f"[DEBUG] Set btn_mute to: {input_id}")
            
            self.save_config()
            print(f"[DEBUG] Config saved: {self.config}")
            
            if self.window:
                try:
                    self.window.evaluate_js(f"updateBinding('{target}', {safe_input})")
                    print(f"[DEBUG] Called updateBinding JS")
                except Exception as e:
                    print(f"Update binding error: {e}")
            
            # Clear pending state after binding is complete
            self.binding_pending = False
            return
        
        # If we're in pending state (API was called but binding_target not set yet), block triggers
        if self.binding_pending:
            print(f"[DEBUG] Blocking trigger - binding_pending is True")
            return
        
        # Normal mode - trigger actions
        if self.running and self.loop and self.rpc_client:
            if input_id == self.config.get('btn_deafen'):
                self.trigger_action('deafen')
            elif input_id == self.config.get('btn_mute'):
                self.trigger_action('mute')

    def trigger_action(self, action_type):
        """Trigger mute/deafen action safely"""
        if not self.running or not self.loop or not self.rpc_client:
            print("[API] Cannot trigger action: RPC not running")
            return
            
        if action_type == 'deafen':
            asyncio.run_coroutine_threadsafe(self._toggle_deaf(), self.loop)
        elif action_type == 'mute':
            asyncio.run_coroutine_threadsafe(self._toggle_mute(), self.loop)
    
    # === UI Update Helpers ===
    
    def update_status(self, message):
        if self.window:
            safe_msg = json.dumps(message)
            try:
                self.window.evaluate_js(f"updateStatus({safe_msg})")
            except:
                pass

    
    def update_connection_status(self, connected):
        if self.window:
            self.window.evaluate_js(f"updateConnectionStatus({'true' if connected else 'false'})")
    
    def update_voice_status(self):
        deaf = self.current_voice_settings['deaf']
        mute = self.current_voice_settings['mute']
        if self.window:
            self.window.evaluate_js(f"updateVoiceStatus({'true' if deaf else 'false'}, {'true' if mute else 'false'})")
    
    # === Tray ===
    
    def create_tray_image(self):
        """Create a beautiful Discord-style tray icon with microphone design"""
        width, height = 64, 64
        
        # Create image with transparency support
        image = Image.new('RGBA', (width, height), (0, 0, 0, 0))
        dc = ImageDraw.Draw(image)
        
        # Draw rounded rectangle background with Discord gradient colors
        # Main background circle - Discord blurple
        dc.ellipse([2, 2, 62, 62], fill=(88, 101, 242, 255))
        
        # Inner subtle highlight ring
        dc.ellipse([4, 4, 60, 60], fill=(99, 112, 250, 255))
        dc.ellipse([6, 6, 58, 58], fill=(88, 101, 242, 255))
        
        # Draw microphone icon (stylized, centered)
        mic_color = (255, 255, 255, 255)
        
        # Microphone body (pill shape)
        dc.rounded_rectangle([26, 14, 38, 34], radius=6, fill=mic_color)
        
        # Microphone stand arc (U shape around the mic)
        dc.arc([20, 22, 44, 44], 0, 180, fill=mic_color, width=3)
        
        # Microphone stand vertical line
        dc.rectangle([30, 42, 34, 48], fill=mic_color)
        
        # Microphone base (horizontal line)
        dc.rounded_rectangle([22, 47, 42, 51], radius=2, fill=mic_color)
        
        # Convert to RGB for pystray compatibility (with Discord color background)
        rgb_image = Image.new('RGB', (width, height), (88, 101, 242))
        rgb_image.paste(image, mask=image.split()[3])
        
        return rgb_image
    
    def run_tray(self):
        image = self.create_tray_image()
        menu = pystray.Menu(
            pystray.MenuItem("開啟", self.show_window, default=True),
            pystray.MenuItem("退出", self.quit_app)
        )
        self.tray_icon = pystray.Icon("DiscordMouseRPC", image, "Discord Mouse Controller", menu)
        self.tray_icon.run()
    
    def show_window(self):
        if self.window:
            self.window.show()
        if self.tray_icon:
            self.tray_icon.stop()
    
    def quit_app(self, icon=None, item=None):
        """Safely quit the application"""
        print("[DEBUG] quit_app called")
        self.running = False
        
        # Stop listeners first
        try:
            if hasattr(self, 'mouse_listener'):
                self.mouse_listener.stop()
            if hasattr(self, 'keyboard_listener'):
                self.keyboard_listener.stop()
        except:
            pass
        
        # Stop tray icon in a separate step
        def stop_tray():
            try:
                if self.tray_icon:
                    self.tray_icon.stop()
            except:
                pass
        
        # Run tray stop in thread to avoid blocking
        if self.tray_icon:
            threading.Thread(target=stop_tray, daemon=True).start()
        
        # Give time for cleanup
        import time
        time.sleep(0.2)
        
        # Force exit
        os._exit(0)


def on_closing(window):
    """Handle window close"""
    api = window._js_api
    if api.config.get('minimize_to_tray'):
        window.hide()
        threading.Thread(target=api.run_tray, daemon=True).start()
        return False  # Prevent close
    else:
        # Clean up resources
        api.running = False
        try:
            if api.mouse_listener:
                api.mouse_listener.stop()
        except:
            pass
        try:
            if api.keyboard_listener:
                api.keyboard_listener.stop()
        except:
            pass
        if api.loop and api.loop.is_running():
            api.loop.call_soon_threadsafe(api.loop.stop)
        return True  # Allow close


def main():
    api = DiscordAPI()
    
    window = webview.create_window(
        title="Discord Mouse Controller",
        url=HTML_FILE,
        width=800,
        height=520,
        resizable=False,
        frameless=True,  # Remove native Windows title bar for macOS-style custom title bar
        easy_drag=True,  # Allow window dragging
        js_api=api,
        background_color='#f5f7fa'
    )
    
    def on_loaded():
        api.set_window(window)
    
    window.events.loaded += on_loaded
    window.events.closing += lambda: on_closing(window)
    
    webview.start(debug=False)


if __name__ == "__main__":
    main()
