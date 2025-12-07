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
import ctypes

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(os.path.abspath(__file__))

    return os.path.join(base_path, relative_path)

# Use absolute path for config file to ensure it works when launched from startup
def get_config_path():
    """Get absolute path to config file, works for both dev and PyInstaller"""
    if getattr(sys, 'frozen', False):
        # When running as exe, use the exe's directory
        base_path = os.path.dirname(sys.executable)
    else:
        # When running as script, use the script's directory
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, "config.json")

CONFIG_FILE = get_config_path()
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
        
        # Combo key tracking - support any two keys
        self.pressed_keys = []  # Track all currently pressed keys (Ordered List)
        self.last_key_time = 0  # Time of last key press for combo detection
        self.combo_timeout = 0.5  # Seconds to consider keys as part of a combo
        
        # Long press tracking for combo binding
        self.first_key_press_time = None  # Time when first key was pressed
        self.long_press_threshold = 0.8  # Seconds to trigger long press binding mode
        self.long_press_timer = None  # Timer thread for long press detection
        self.pending_combo = None  # The combo being built during long press
        self.long_press_active = False  # Flag to indicate long press mode is active
        
        # Config
        self.config = self.load_config()
        self.saved_access_token = self.config.get('access_token')
        self.saved_refresh_token = self.config.get('refresh_token')
        
        # Start input listeners
        self.mouse_listener = mouse.Listener(on_click=self.on_click)
        self.mouse_listener.start()
        
        self.keyboard_listener = keyboard.Listener(
            on_press=self.on_key_press,
            on_release=self.on_key_release
        )
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
    
    def has_config(self):
        """Check if config file exists - used for onboarding detection"""
        return os.path.exists(CONFIG_FILE)
    
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
        """Close window - handle minimize to tray or actual close"""
        print(f"[DEBUG] close_window called. minimize_to_tray={self.config.get('minimize_to_tray')}")
        
        if self.config.get('minimize_to_tray'):
            # Minimize to tray instead of closing
            print("[DEBUG] Minimizing to tray...")
            if self.window:
                # Start tray first, then hide window
                def do_minimize_to_tray():
                    try:
                        self.run_tray()
                    except Exception as e:
                        print(f"[ERROR] Tray error: {e}")
                
                tray_thread = threading.Thread(target=do_minimize_to_tray, daemon=True)
                tray_thread.start()
                
                # Give tray a moment to start, then hide
                time.sleep(0.1)
                self.window.hide()
                print("[DEBUG] Window hidden, tray should be running")
        else:
            # Actually close the application
            print("[DEBUG] Actually closing...")
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
        """Set or unset app startup in Windows Registry with --minimized flag"""
        key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
        app_name = "DiscordMouseController"
        
        # Build exe path with --minimized flag for startup
        if getattr(sys, 'frozen', False):
            exe_path = f'"{sys.executable}" --minimized'
        else:
            exe_path = f'"{sys.executable}" "{os.path.abspath(sys.argv[0])}" --minimized'

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
    
    def toggle_media(self):
        """Toggle media play/pause using Windows API"""
        print("[API] Toggle Media requested")
        self.trigger_action('media')
        return True
    
    def send_media_key(self):
        """Send media play/pause key using Windows API"""
        try:
            # VK_MEDIA_PLAY_PAUSE = 0xB3
            VK_MEDIA_PLAY_PAUSE = 0xB3
            KEYEVENTF_EXTENDEDKEY = 0x0001
            KEYEVENTF_KEYUP = 0x0002
            
            # Key down
            ctypes.windll.user32.keybd_event(VK_MEDIA_PLAY_PAUSE, 0, KEYEVENTF_EXTENDEDKEY, 0)
            time.sleep(0.05)
            # Key up
            ctypes.windll.user32.keybd_event(VK_MEDIA_PLAY_PAUSE, 0, KEYEVENTF_EXTENDEDKEY | KEYEVENTF_KEYUP, 0)
            print("[DEBUG] Media key sent successfully")
        except Exception as e:
            print(f"[ERROR] Failed to send media key: {e}")
    
    def start_drag(self):
        """Start window drag - for custom title bar"""
        if self.window:
            self.window.move(0, 0)  # This enables native drag
    
    def set_bind_target(self, target_type):
        """Set the binding target - 0 for deafen, 1 for mute, 2 for media"""
        print(f"[API] set_bind_target called with: {target_type}")
        
        # Immediately block action triggers
        self.binding_pending = True
        
        if target_type == 0:
            self.start_binding('deafen')
        elif target_type == 1:
            self.start_binding('mute')
        elif target_type == 2:
            self.start_binding('media')
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
        elif target == 'media':
            print("[-DEBUG-] Target MATCHES 'media'")
        else:
            print(f"[-DEBUG-] Target '{target}' DOES NOT MATCH expected values!")

        self.binding_target = target
        print(f"[-DEBUG-] binding_target NOW SET TO: '{self.binding_target}'")
        
        # Explicit feedback to UI
        target_names = {'deafen': '拒聽', 'mute': '靜音', 'media': '媒體'}
        target_name = target_names.get(target, target)
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
                    # Use verify async sleep to keep loop running!
                    self.loop.run_until_complete(asyncio.sleep(2))
            except Exception as e:
                print(f"RPC Error: {e}")
                if self.running:
                    try:
                        self.loop.run_until_complete(asyncio.sleep(2))
                    except:
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
                # Token exchange with retry
                max_auth_attempts = 3
                for auth_attempt in range(max_auth_attempts):
                    try:
                        self.update_status(f"請在 Discord 視窗中點擊授權... (嘗試 {auth_attempt + 1}/{max_auth_attempts})")
                        auth_resp = await self.rpc_client.authorize(client_id, scopes=['rpc'])
                        code = auth_resp['data']['code']
                        
                        # Exchange immediately to avoid code expiration
                        self.update_status("交換 Token 中...")
                        
                        # Debug: print what we're sending (hide secret)
                        print(f"[DEBUG] Token exchange - client_id: {client_id[:8]}..., secret_len: {len(client_secret)}")
                        
                        token_resp = requests.post(
                            'https://discord.com/api/oauth2/token',
                            data={
                                'client_id': client_id,
                                'client_secret': client_secret,
                                'grant_type': 'authorization_code',
                                'code': code,
                                'redirect_uri': 'http://127.0.0.1'
                            },
                            headers={
                                'Content-Type': 'application/x-www-form-urlencoded'
                            },
                            timeout=10
                        )
                        
                        if token_resp.status_code == 200:
                            token_data = token_resp.json()
                            access_token = token_data['access_token']
                            refresh_token = token_data.get('refresh_token')
                            
                            self.saved_access_token = access_token
                            self.saved_refresh_token = refresh_token
                            self.save_config()
                            break  # Success, exit retry loop
                        else:
                            error_text = token_resp.text
                            print(f"Token exchange attempt {auth_attempt + 1} failed: {error_text}")
                            
                            # Check for specific errors
                            if 'invalid_client' in error_text:
                                if auth_attempt >= max_auth_attempts - 1:
                                    raise Exception("Client Secret 不正確，請檢查 Discord Developer Portal 中的設定")
                            
                            if auth_attempt < max_auth_attempts - 1:
                                self.update_status(f"Token 交換失敗，重試中... ({auth_attempt + 2}/{max_auth_attempts})")
                                await asyncio.sleep(1)
                            else:
                                raise Exception(f"Token 交換失敗: {error_text}")
                    except Exception as auth_error:
                        if auth_attempt < max_auth_attempts - 1:
                            print(f"Auth attempt {auth_attempt + 1} failed: {auth_error}")
                            await asyncio.sleep(1)
                        else:
                            raise
            
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
            # Don't set running=False here - let the loop retry
            # Only update UI if window is visible
            if self.window:
                try:
                    self.update_status(f"連線中斷，重新連接中...")
                    self.update_connection_status(False)
                except:
                    pass
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
    
    def _normalize_key(self, key_str):
        """Normalize key names for cleaner display"""
        # Remove Key. prefix and clean up
        clean = key_str.replace("Key.", "").replace("'", "")
        
        # Normalize common keys
        normalizations = {
            'ctrl_l': 'Ctrl', 'ctrl_r': 'Ctrl', 'ctrl': 'Ctrl',
            'shift_l': 'Shift', 'shift_r': 'Shift', 'shift': 'Shift',
            'alt_l': 'Alt', 'alt_r': 'Alt', 'alt': 'Alt', 'alt_gr': 'Alt',
            'cmd': 'Win', 'cmd_l': 'Win', 'cmd_r': 'Win',
            'space': 'Space', 'enter': 'Enter', 'esc': 'Esc',
            'tab': 'Tab', 'backspace': 'Backspace',
            'up': '↑', 'down': '↓', 'left': '←', 'right': '→',
            'f1': 'F1', 'f2': 'F2', 'f3': 'F3', 'f4': 'F4',
            'f5': 'F5', 'f6': 'F6', 'f7': 'F7', 'f8': 'F8',
            'f9': 'F9', 'f10': 'F10', 'f11': 'F11', 'f12': 'F12',
        }
        
        lower_clean = clean.lower()
        if lower_clean in normalizations:
            return normalizations[lower_clean]
        
        # For mouse buttons
        if 'Button.' in key_str:
            button_name = key_str.replace('Button.', '')
            button_map = {'left': 'LMB', 'right': 'RMB', 'middle': 'MMB', 
                         'x1': 'Mouse4', 'x2': 'Mouse5'}
            return button_map.get(button_name, button_name)
        
        # Return uppercase single char or as-is
        if len(clean) == 1:
            return clean.upper()
        return clean
    
    def on_key_press(self, key):
        key_str = str(key)
        normalized = self._normalize_key(key_str)
        current_time = time.time()
        
        # If in binding mode
        if self.binding_target:
            # Add to pressed keys if not present
            if normalized not in self.pressed_keys:
                self.pressed_keys.append(normalized)
            
            # If this is the first key pressed, start long press timer
            if len(self.pressed_keys) == 1 and self.first_key_press_time is None:
                self.first_key_press_time = current_time
                self.pending_combo = normalized
                self.long_press_active = False
                
                # Start a timer to detect long press
                def long_press_check():
                    time.sleep(self.long_press_threshold)
                    # Check if still pressing the same key
                    if self.binding_target and normalized in self.pressed_keys:
                        self.long_press_active = True
                        self.pending_combo = normalized
                        if self.window:
                            try:
                                self.window.evaluate_js(f"updateStatus('已鎖定 {normalized}，請按第二個按鍵...')")
                            except:
                                pass
                        print(f"[DEBUG] Long press detected: {normalized}, waiting for second key...")
                
                self.long_press_timer = threading.Thread(target=long_press_check, daemon=True)
                self.long_press_timer.start()
                
            elif len(self.pressed_keys) >= 2 and self.long_press_active:
                # Second key pressed during long press mode - create combo
                # Build combo with both keys - using list order
                combo = self._build_combo_string_from_list(self.pressed_keys)
                self.pending_combo = combo
                print(f"[DEBUG] Combo key detected: {combo}")
                if self.window:
                    try:
                        self.window.evaluate_js(f"updateStatus('組合鍵: {combo}')")
                    except:
                        pass
            return
        
        # Normal mode - add to pressed keys and track for release
        if normalized not in self.pressed_keys:
            self.pressed_keys.append(normalized)
        self.last_key_time = current_time
    
    def on_key_release(self, key):
        key_str = str(key)
        normalized = self._normalize_key(key_str)
        
        # If in binding mode and a key was released
        if self.binding_target:
            # ESC cancels binding
            if normalized == 'Esc':
                self._cancel_binding()
                self.pressed_keys.clear()
                self._reset_long_press_state()
                return
            
            # When all keys are released, complete the binding
            if normalized in self.pressed_keys:
                self.pressed_keys.remove(normalized)
                
            if len(self.pressed_keys) == 0 and self.pending_combo:
                # Complete binding with the pending combo
                self._complete_binding(self.pending_combo)
                self._reset_long_press_state()
            return
        
        # Normal mode - Check match on release
        # We check the combo formed by (keys still held) + (key being released)
        # This allows "Hold A, Press B, Release B" to trigger "A+B"
        
        if normalized not in self.pressed_keys:
            self.pressed_keys.append(normalized)
        
        # Build combo string for this moment
        # Note: pressed_keys is now a list, so order is preserved
        combo_to_check = self._build_combo_string_from_list(self.pressed_keys)
        
        # Check against config and trigger
        triggered = self._check_and_trigger(combo_to_check)
        
        # Remove from pressed keys
        if normalized in self.pressed_keys:
            self.pressed_keys.remove(normalized)
    
    def _build_combo_string_from_list(self, key_list):
        """Build combo string from a list of keys, preserving order for non-modifiers"""
        if not key_list:
            return ""
        
        # Priority: Ctrl, Alt, Shift, Win
        priority = {'Ctrl': 0, 'Alt': 1, 'Shift': 2, 'Win': 3}
        
        modifiers = []
        others = []
        
        # Separate modifiers and others
        # We process the input list in order
        for k in key_list:
            if k in priority:
                modifiers.append(k)
            else:
                others.append(k)
        
        # Sort modifiers by priority
        modifiers.sort(key=lambda x: priority[x])
        
        # Combine: Modifiers first, then others (in original press order)
        final_list = modifiers + others
        
        return '+'.join(final_list[:2])

    def _build_combo_string(self):
        """Build a combo string from currently pressed keys (max 2 keys)"""
        return self._build_combo_string_from_list(self.pressed_keys)
        
    def _check_and_trigger(self, combo_id):
        """Check if a combo is bound and trigger it"""
        if not combo_id:
            return False
            
        # Update UI with last input for visual feedback
        if self.window:
            try:
                # Only update simple inputs or if actually bound to avoid spamming UI
                # But for debugging, showing everything is helpful
                safe_input = json.dumps(combo_id)
                self.window.evaluate_js(f"updateLastInput({safe_input})")
            except:
                pass
                
        # Check bindings
        triggered = False
        if combo_id == self.config.get('btn_media'):
            self.trigger_action('media')
            triggered = True
        
        if combo_id == self.config.get('btn_deafen'):
            self.trigger_action('deafen')
            triggered = True
        elif combo_id == self.config.get('btn_mute'):
            self.trigger_action('mute')
            triggered = True
            
        return triggered

    def _reset_long_press_state(self):
        """Reset all long press related state"""
        self.first_key_press_time = None
        self.pending_combo = None
        self.long_press_active = False
        self.long_press_timer = None
    
    def _complete_binding(self, input_id):
        """Complete the key binding process"""
        target = self.binding_target
        self.binding_target = None
        self.binding_pending = False
        
        safe_input = json.dumps(input_id)
        
        print(f"[DEBUG] Completing binding: {target} = {input_id}")
        
        if target == 'deafen':
            self.config['btn_deafen'] = input_id
        elif target == 'mute':
            self.config['btn_mute'] = input_id
        elif target == 'media':
            self.config['btn_media'] = input_id
        
        self.save_config()
        
        if self.window:
            try:
                self.window.evaluate_js(f"updateBinding('{target}', {safe_input})")
            except Exception as e:
                print(f"Update binding error: {e}")
    
    def _cancel_binding(self):
        """Cancel the current binding and clear it"""
        target = self.binding_target
        self.binding_target = None
        self.binding_pending = False
        
        print(f"[DEBUG] ESC pressed - clearing binding for {target}")
        
        if target == 'deafen':
            self.config['btn_deafen'] = None
        elif target == 'mute':
            self.config['btn_mute'] = None
        elif target == 'media':
            self.config['btn_media'] = None
        
        self.save_config()
        
        if self.window:
            try:
                self.window.evaluate_js(f"updateBinding('{target}', 'None')")
                self.window.evaluate_js(f"showNotification('已取消 {target} 綁定')")
            except Exception as e:
                print(f"Update binding error: {e}")
    
    def on_click(self, x, y, button, pressed):
        button_str = str(button)
        normalized = self._normalize_key(button_str)
        
        current_time = time.time()
        
        if pressed:
            # If in binding mode
            if self.binding_target:
                # Ignore left click during binding (to prevent accidental clicks)
                if normalized == "LMB":
                    return
                
                if normalized not in self.pressed_keys:
                    self.pressed_keys.append(normalized)
                
                # If this is the first key pressed, start long press timer
                if len(self.pressed_keys) == 1 and self.first_key_press_time is None:
                    self.first_key_press_time = current_time
                    self.pending_combo = normalized
                    self.long_press_active = False
                    
                    # Start a timer to detect long press
                    def long_press_check():
                        time.sleep(self.long_press_threshold)
                        # Check if still pressing the same key
                        if self.binding_target and normalized in self.pressed_keys:
                            self.long_press_active = True
                            self.pending_combo = normalized
                            if self.window:
                                try:
                                    self.window.evaluate_js(f"updateStatus('已鎖定 {normalized}，請按第二個按鍵...')")
                                except:
                                    pass
                            print(f"[DEBUG] Long press detected: {normalized}, waiting for second key...")
                    
                    self.long_press_timer = threading.Thread(target=long_press_check, daemon=True)
                    self.long_press_timer.start()
                    
                elif len(self.pressed_keys) >= 2 and self.long_press_active:
                    # Second key pressed during long press mode - create combo
                    combo = self._build_combo_string_from_list(self.pressed_keys)
                    self.pending_combo = combo
                    print(f"[DEBUG] Combo key detected: {combo}")
                    if self.window:
                        try:
                            self.window.evaluate_js(f"updateStatus('組合鍵: {combo}')")
                        except:
                            pass
                return
            
            # Normal mode - track pressed keys
            if normalized not in self.pressed_keys:
                self.pressed_keys.append(normalized)
            
        else:  # Released
            # If in binding mode
            if self.binding_target:
                if normalized in self.pressed_keys:
                    self.pressed_keys.remove(normalized)
                if len(self.pressed_keys) == 0 and self.pending_combo:
                    # Complete binding with the pending combo
                    self._complete_binding(self.pending_combo)
                    self._reset_long_press_state()
                return
            
            # Normal mode - trigger action on release
            current_combo_keys = self.pressed_keys.copy()
            if normalized not in current_combo_keys:
                current_combo_keys.append(normalized) # Ensure releasing key is included
            
            combo_to_check = self._build_combo_string_from_list(current_combo_keys)
            
            # Check matches
            self._check_and_trigger(combo_to_check)
            
            # Cleanup
            if normalized in self.pressed_keys:
                self.pressed_keys.remove(normalized)
    
    def _build_combo_string(self):
        """Build a combo string from currently pressed keys (max 2 keys)"""
        if not self.pressed_keys:
            return ""
        
        # Sort keys for consistent ordering
        # Priority: Ctrl, Alt, Shift, Win, then alphabetical
        priority = {'Ctrl': 0, 'Alt': 1, 'Shift': 2, 'Win': 3}
        sorted_keys = sorted(self.pressed_keys, 
                           key=lambda x: (priority.get(x, 10), x))
        
        # Take max 2 keys
        keys_to_use = sorted_keys[:2]
        return '+'.join(keys_to_use)
    
    def _handle_input(self, input_id):
        """Handle input in normal mode - trigger bound actions"""
        print(f"[DEBUG] _handle_input called with: {input_id}")
        
        # Escape for safe JS injection
        safe_input = json.dumps(input_id)
        
        # Update UI with last input
        if self.window:
            try:
                self.window.evaluate_js(f"updateLastInput({safe_input})")
            except:
                pass
        
        # If in binding mode, ignore (handled elsewhere)
        if self.binding_target or self.binding_pending:
            return
        
        # Normal mode - trigger actions
        # Media can work without RPC connection
        if input_id == self.config.get('btn_media'):
            self.trigger_action('media')
        
        # Discord actions need RPC
        if self.loop and self.rpc_client:
            if input_id == self.config.get('btn_deafen'):
                self.trigger_action('deafen')
            elif input_id == self.config.get('btn_mute'):
                self.trigger_action('mute')

    def trigger_action(self, action_type):
        """Trigger mute/deafen/media action safely"""
        print(f"[DEBUG] trigger_action called: {action_type}")
        print(f"[DEBUG] running={self.running}, loop={self.loop is not None}, rpc={self.rpc_client is not None}")
        
        # Media action doesn't need RPC
        if action_type == 'media':
            threading.Thread(target=self.send_media_key, daemon=True).start()
            return
        
        if not self.rpc_client:
            print("[API] Cannot trigger action: RPC not connected")
            return
        
        # Try to use existing loop if running
        if self.loop and self.loop.is_running():
            try:
                if action_type == 'deafen':
                    future = asyncio.run_coroutine_threadsafe(self._toggle_deaf(), self.loop)
                elif action_type == 'mute':
                    future = asyncio.run_coroutine_threadsafe(self._toggle_mute(), self.loop)
                else:
                    return
                
                # Add error logging callback
                def log_error(fut):
                    try:
                        fut.result(timeout=1)
                    except Exception as e:
                        print(f"[ERROR] Action {action_type} failed: {e}")
                        # Fallback to sync send
                        self._sync_trigger_action(action_type)
                        
                future.add_done_callback(log_error)
                return
            except Exception as e:
                print(f"[DEBUG] Async trigger failed: {e}, trying sync fallback")
        
        # Fallback: use sync method
        self._sync_trigger_action(action_type)
    
    def _sync_trigger_action(self, action_type):
        """Synchronous fallback for triggering actions when event loop isn't available"""
        print(f"[DEBUG] _sync_trigger_action called: {action_type}")
        
        if not self.rpc_client or not hasattr(self.rpc_client, 'sock_writer'):
            print("[API] Cannot sync trigger: RPC socket not available")
            return
            
        try:
            if action_type == 'deafen':
                new_value = not self.current_voice_settings.get('deaf', False)
                payload = {
                    'cmd': 'SET_VOICE_SETTINGS',
                    'args': {'deaf': new_value},
                    'nonce': str(time.time())
                }
            elif action_type == 'mute':
                new_value = not self.current_voice_settings.get('mute', False)
                payload = {
                    'cmd': 'SET_VOICE_SETTINGS',
                    'args': {'mute': new_value},
                    'nonce': str(time.time())
                }
            else:
                return
            
            # Send synchronously
            payload_json = json.dumps(payload)
            encoded = payload_json.encode('utf-8')
            header = struct.pack('<II', 1, len(encoded))
            
            # Use a thread to avoid blocking
            def do_send():
                try:
                    # Access the underlying socket directly
                    sock_writer = self.rpc_client.sock_writer
                    sock_writer.write(header + encoded)
                    # Note: drain() is async, but for fire-and-forget this should work
                    print(f"[DEBUG] Sync send completed for {action_type}")
                except Exception as e:
                    print(f"[ERROR] Sync send failed: {e}")
            
            threading.Thread(target=do_send, daemon=True).start()
            
        except Exception as e:
            print(f"[ERROR] _sync_trigger_action failed: {e}")
    
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
        # Prevent duplicate tray icons
        if self.tray_icon:
            try:
                self.tray_icon.stop()
            except:
                pass
        
        print("[DEBUG] Starting tray icon...")
        try:
            image = self.create_tray_image()
            menu = pystray.Menu(
                pystray.MenuItem("開啟", self.show_window, default=True),
                pystray.MenuItem("退出", self.quit_app)
            )
            self.tray_icon = pystray.Icon("DiscordMouseRPC", image, "Discord Mouse Controller", menu)
            # Use threading to be absolutely sure we don't block anything
            self.tray_icon.run() 
        except Exception as e:
            print(f"[ERROR] Failed to start tray icon: {e}")

    def show_window(self):
        """Restore the window from tray"""
        if self.window:
            self.window.show()
            self.window.restore()  # Ensure it's not minimized
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
        
        # Stop tray icon
        try:
            if self.tray_icon:
                self.tray_icon.stop()
        except:
            pass
        
        # Cleanup threading
        import time
        time.sleep(0.2)
        
        # Force exit
        os._exit(0)


def on_closing(window):
    """Handle window close"""
    api = window._js_api
    print(f"[DEBUG] on_closing called. minimize_to_tray={api.config.get('minimize_to_tray')}")
    
    if api.config.get('minimize_to_tray'):
        # Minimize to tray instead of closing
        # We need to prevent the actual close and just hide the window
        print("[DEBUG] Minimize to tray is enabled")
        
        # Schedule hide and tray operations
        def do_minimize_to_tray():
            try:
                # First start the tray in a separate thread (non-blocking)
                tray_thread = threading.Thread(target=api.run_tray, daemon=True)
                tray_thread.start()
                
                # Small delay to let tray initialize
                time.sleep(0.1)
                
                # Then hide the window
                if window:
                    window.hide()
                    
                print("[DEBUG] Window hidden, tray should be running")
            except Exception as e:
                print(f"[ERROR] do_minimize_to_tray failed: {e}")
        
        # Run in a thread to avoid blocking
        threading.Thread(target=do_minimize_to_tray, daemon=True).start()
        
        # Return False to prevent window destruction
        # Note: In some pywebview backends this may not work perfectly
        return False
    else:
        # Clean up resources only when actually closing
        print("[DEBUG] Actually closing the app...")
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
    # Check if launched with --minimized flag (for startup)
    start_minimized = '--minimized' in sys.argv
    
    if start_minimized:
        print("[DEBUG] Starting in minimized mode (startup)")
    
    api = DiscordAPI()
    
    # If starting minimized, we need to set minimize_to_tray to true
    # and start hidden directly into system tray
    if start_minimized:
        api.config['minimize_to_tray'] = True
        api.save_config()
    
    window = webview.create_window(
        title="Discord Mouse Controller",
        url=HTML_FILE,
        width=800,
        height=520,
        resizable=False,
        frameless=True,  # Remove native Windows title bar for macOS-style custom title bar
        easy_drag=True,  # Allow window dragging
        js_api=api,
        background_color='#f5f7fa',
        hidden=start_minimized  # Start hidden if minimized mode
    )
    
    def on_loaded():
        api.set_window(window)
        
        # If started minimized, immediately go to tray
        if start_minimized:
            def start_in_tray():
                time.sleep(0.5)  # Give window time to fully initialize
                # Start tray icon
                tray_thread = threading.Thread(target=api.run_tray, daemon=True)
                tray_thread.start()
            threading.Thread(target=start_in_tray, daemon=True).start()
    
    window.events.loaded += on_loaded
    
    def handle_closing():
        return on_closing(window)
    
    window.events.closing += handle_closing
    
    webview.start(debug=False)


if __name__ == "__main__":
    main()
