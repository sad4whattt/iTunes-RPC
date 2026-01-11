import time
import threading
import win32com.client
import pythoncom
import requests
from pypresence import Presence
import urllib.parse
import pystray
from PIL import Image, ImageDraw

CLIENT_ID = '1459800355243163846'
FALLBACK_IMAGE = 'itunes_logo'

class RPCHandler:
    def __init__(self):
        self.rpc = Presence(CLIENT_ID)
        self.itunes = None
        self.last_track_id = None
        self.cached_artwork_url = FALLBACK_IMAGE
        self.current_track_info = None
        self.rpc_enabled = True
        self.running = True
        self.play_counts = {}
        self.track_started_at = None
        self.track_counted = False

    def connect(self):
        try:
            self.rpc.connect()
            print("Connected to Discord.")
        except Exception as e:
            print(f"Error connecting to Discord: {e}")

    def clean_string(self, s):
        if not s: return ""
        return ''.join(e for e in s.lower() if e.isalnum())

    def fetch_artwork_url(self, artist, album, song_name):
        if not artist or not song_name:
            return FALLBACK_IMAGE

        def search_apple_music(search_term, entity_type):
            try:
                encoded_query = urllib.parse.quote(search_term)
                url = f"https://itunes.apple.com/search?term={encoded_query}&media=music&entity={entity_type}&limit=5"
                response = requests.get(url, timeout=5)
                data = response.json()

                if data['resultCount'] == 0: return None
                target_artist = self.clean_string(artist)
                
                for result in data['results']:
                    api_artist = self.clean_string(result.get('artistName', ''))
                    if target_artist in api_artist or api_artist in target_artist:
                        return result['artworkUrl100'].replace('100x100bb', '600x600bb')
                return None
            except Exception:
                return None

        image = search_apple_music(f"{artist} {song_name}", "musicTrack")
        if image: return image
        
        if album:
            image = search_apple_music(f"{artist} {album}", "album")
            if image: return image

        return FALLBACK_IMAGE

    def get_track_info(self):
        try:
            if self.itunes is None:
                self.itunes = win32com.client.Dispatch("iTunes.Application")

            if self.itunes.PlayerState != 1:
                return None
            
            track = self.itunes.CurrentTrack
            unique_id = f"{track.Name}-{track.Artist}-{track.Album}"

            return {
                "id": unique_id,
                "name": track.Name,
                "artist": track.Artist,
                "album": track.Album,
                "position": self.itunes.PlayerPosition,
                "duration": track.Duration
            }
        except Exception:
            self.itunes = None
            return None

    def toggle_rpc(self):
        """Called by the tray menu to toggle status"""
        self.rpc_enabled = not self.rpc_enabled
        if not self.rpc_enabled:
            print("RPC Disabled by user. Clearing status.")
            self.rpc.clear()
            self.last_track_id = None

    def loop(self):
        pythoncom.CoInitialize()
        while self.running:
            try:
                if not self.rpc_enabled:
                    time.sleep(2)
                    continue

                track = self.get_track_info()
                self.current_track_info = track

                if track:
                    if track['id'] != self.last_track_id:
                        self.cached_artwork_url = self.fetch_artwork_url(track['artist'], track['album'], track['name'])
                        self.last_track_id = track['id']
                        self.track_started_at = time.time()
                        self.track_counted = False

                    if not self.track_counted and self.track_started_at:
                        time_listened = time.time() - self.track_started_at
                        threshold = min(30, track['duration'] * 0.5)
                        if time_listened >= threshold:
                            self.play_counts[track['id']] = self.play_counts.get(track['id'], 0) + 1
                            self.track_counted = True

                    time_remaining = track['duration'] - track['position']
                    end_timestamp = int(time.time() + time_remaining)
                    start_timestamp = int(time.time() - track['position'])

                    play_count = self.play_counts.get(track['id'], 0)
                    play_text = f"{play_count} play{'s' if play_count != 1 else ''} this session"

                    self.rpc.update(
                        details=track['name'],
                        state=f"by {track['artist']} • {play_text}",
                        large_image=self.cached_artwork_url,
                        large_text=track['album'],
                        start=start_timestamp,
                        end=end_timestamp
                    )
                else:
                    self.rpc.clear()
                    self.last_track_id = None
                    self.track_started_at = None
                    self.track_counted = False

                time.sleep(3)

            except Exception as e:
                print(f"Loop error: {e}")
                time.sleep(3)


handler = RPCHandler()

def get_menu_items():
    """Generates the menu items dynamically every time you right-click."""
    
    items = []
    
    if handler.current_track_info and handler.rpc_enabled:
        t = handler.current_track_info
        items.append(pystray.MenuItem(f"🎵 {t['name']}", lambda i, m: None, enabled=False))
        items.append(pystray.MenuItem(f"👤 {t['artist']}", lambda i, m: None, enabled=False))
        items.append(pystray.MenuItem(f"💿 {t['album']}", lambda i, m: None, enabled=False))
        items.append(pystray.Menu.SEPARATOR)
    elif not handler.rpc_enabled:
        items.append(pystray.MenuItem("⚠️ RPC is Paused", lambda i, m: None, enabled=False))
        items.append(pystray.Menu.SEPARATOR)
    else:
        items.append(pystray.MenuItem("No Music Playing", lambda i, m: None, enabled=False))
        items.append(pystray.Menu.SEPARATOR)

    toggle_label = "Enable RPC" if not handler.rpc_enabled else "Disable RPC"
    
    def on_toggle(icon, item):
        handler.toggle_rpc()

    items.append(pystray.MenuItem(toggle_label, on_toggle))
    items.append(pystray.MenuItem("Quit", quit_action))
    
    return items

def create_image():
    width = 64
    height = 64
    color1 = "black"
    color2 = "white"
    image = Image.new('RGB', (width, height), color1)
    dc = ImageDraw.Draw(image)
    dc.rectangle((width // 2, 0, width, height // 2), fill=color2)
    dc.rectangle((0, height // 2, width // 2, height), fill=color2)
    return image

def quit_action(icon, item):
    handler.running = False
    icon.stop()

def run_background_rpc():
    handler.connect()
    handler.loop()

if __name__ == '__main__':
    rpc_thread = threading.Thread(target=run_background_rpc)
    rpc_thread.start()
    icon = pystray.Icon("iTunesRPC", create_image(), "iTunes RPC", menu=pystray.Menu(get_menu_items))

    icon.run()