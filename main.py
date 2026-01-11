import json
import time
import threading
import queue
import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path
import urllib.parse

import win32com.client
import pythoncom
import requests
from pypresence import Presence
import pystray
from PIL import Image, ImageDraw

DEFAULT_CLIENT_ID = "1459800355243163846"
FALLBACK_IMAGE = "itunes_logo"
ROOT_PATH = Path(__file__).resolve().parent
CONFIG_PATH = ROOT_PATH / "config.json"
ARTWORK_CACHE_PATH = ROOT_PATH / "artwork_cache.json"
LOG_PATH = ROOT_PATH / "app.log"


DEFAULT_CONFIG = {
    "client_id": DEFAULT_CLIENT_ID,
    "refresh_interval": 3,
    "idle_interval": 2,
    "request_timeout": 5,
    "max_retry": 3,
    "retry_backoff": 1.5,
    "dry_run": False,
    "privacy": {
        "hide_metadata": False,
        "network_artwork": True,
    },
    "presence_format": {
        "details": "{name}",
        "state": "by {artist} • {play_text}",
    },
    "play_count_threshold_seconds": 30,
    "play_count_threshold_fraction": 0.5,
    "buttons_enabled": True,
    "buttons": [
        {
            "label": "Open in Apple Music",
            "url_template": "https://music.apple.com/search?term={artist}%20{album}%20{name}",
        },
        {
            "label": "View Artist",
            "url_template": "https://music.apple.com/search?term={artist}",
        },
    ],
}


def load_config():
    if CONFIG_PATH.exists():
        try:
            user_cfg = json.loads(CONFIG_PATH.read_text())
            merged = DEFAULT_CONFIG.copy()
            merged.update(user_cfg)
            merged["privacy"] = {**DEFAULT_CONFIG["privacy"], **user_cfg.get("privacy", {})}
            merged["presence_format"] = {
                **DEFAULT_CONFIG["presence_format"],
                **user_cfg.get("presence_format", {}),
            }
            merged["buttons"] = user_cfg.get("buttons", DEFAULT_CONFIG["buttons"])
            return merged
        except Exception:
            pass
    CONFIG_PATH.write_text(json.dumps(DEFAULT_CONFIG, indent=2))
    return DEFAULT_CONFIG


def configure_logging():
    handler = RotatingFileHandler(LOG_PATH, maxBytes=500_000, backupCount=2)
    fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")
    handler.setFormatter(fmt)
    logging.basicConfig(level=logging.INFO, handlers=[handler])
    logging.getLogger().addHandler(logging.StreamHandler())


class RPCHandler:
    def __init__(self, config):
        self.config = config
        self.rpc = Presence(self.config["client_id"])
        self.itunes = None
        self.last_track_id = None
        self.cached_artwork_url = FALLBACK_IMAGE
        self.current_track_info = None
        self.rpc_enabled = True
        self.running = True
        self.play_counts = {}
        self.track_started_at = None
        self.track_counted = False
        self.artwork_cache = self._load_artwork_cache()
        self.artwork_queue = queue.Queue()
        self.artwork_thread = threading.Thread(target=self._artwork_worker, daemon=True)
        self.artwork_thread.start()
        self.menu_track_info = None
        self.menu_track_seen_at = 0

    def _log(self, level, msg):
        logging.log(level, msg)

    def _load_artwork_cache(self):
        if ARTWORK_CACHE_PATH.exists():
            try:
                return json.loads(ARTWORK_CACHE_PATH.read_text())
            except Exception:
                return {}
        return {}

    def _save_artwork_cache(self):
        try:
            ARTWORK_CACHE_PATH.write_text(json.dumps(self.artwork_cache, indent=2))
        except Exception:
            pass

    def connect(self):
        try:
            if self.config.get("dry_run"):
                self._log(logging.INFO, "Dry-run enabled; skipping Discord RPC connect")
                return
            self.rpc.connect()
            self._log(logging.INFO, "Connected to Discord.")
        except Exception as e:
            self._log(logging.WARNING, f"Error connecting to Discord: {e}")

    def ensure_connected(self):
        try:
            if self.config.get("dry_run"):
                return True
            self.rpc.connect()
            return True
        except Exception as e:
            self._log(logging.WARNING, f"Reconnection attempt failed: {e}")
            return False

    def clean_string(self, s):
        if not s:
            return ""
        return "".join(e for e in s.lower() if e.isalnum())

    def fetch_artwork_url(self, artist, album, song_name):
        if not artist or not song_name:
            return FALLBACK_IMAGE

        cache_key = f"{artist}-{album}-{song_name}"
        if cache_key in self.artwork_cache:
            return self.artwork_cache[cache_key]

        def request_json(url):
            delay = 0
            for attempt in range(self.config["max_retry"]):
                if attempt:
                    time.sleep(delay or self.config["retry_backoff"])
                    delay = max(delay * self.config["retry_backoff"], self.config["retry_backoff"])
                try:
                    resp = requests.get(url, timeout=self.config["request_timeout"])
                    resp.raise_for_status()
                    return resp.json()
                except Exception as exc:
                    self._log(logging.DEBUG, f"Artwork request failed (attempt {attempt + 1}): {exc}")
            return None

        def search_apple_music(search_term, entity_type):
            if not self.config["privacy"].get("network_artwork", True):
                return None
            encoded_query = urllib.parse.quote(search_term)
            url = f"https://itunes.apple.com/search?term={encoded_query}&media=music&entity={entity_type}&limit=5"
            data = request_json(url)
            if not data or data.get("resultCount", 0) == 0:
                return None
            target_artist = self.clean_string(artist)
            for result in data.get("results", []):
                api_artist = self.clean_string(result.get("artistName", ""))
                if target_artist in api_artist or api_artist in target_artist:
                    return result.get("artworkUrl100", "").replace("100x100bb", "600x600bb")
            return None

        image = search_apple_music(f"{artist} {song_name}", "musicTrack")
        if image:
            self.artwork_cache[cache_key] = image
            self._save_artwork_cache()
            return image

        if album:
            image = search_apple_music(f"{artist} {album}", "album")
            if image:
                self.artwork_cache[cache_key] = image
                self._save_artwork_cache()
                return image

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
                "duration": track.Duration,
            }
        except Exception as exc:
            self._log(logging.DEBUG, f"Failed to read track info: {exc}")
            self.itunes = None
            return None

    def toggle_rpc(self):
        self.rpc_enabled = not self.rpc_enabled
        if not self.rpc_enabled:
            self._log(logging.INFO, "RPC Disabled by user. Clearing status.")
            self._clear_presence()
            self.last_track_id = None

    def toggle_privacy(self):
        current = self.config["privacy"]["hide_metadata"]
        self.config["privacy"]["hide_metadata"] = not current
        CONFIG_PATH.write_text(json.dumps(self.config, indent=2))
        self._log(logging.INFO, f"Privacy hide_metadata set to {self.config['privacy']['hide_metadata']}")

    def toggle_network_artwork(self):
        current = self.config["privacy"]["network_artwork"]
        self.config["privacy"]["network_artwork"] = not current
        CONFIG_PATH.write_text(json.dumps(self.config, indent=2))
        self._log(logging.INFO, f"Network artwork set to {self.config['privacy']['network_artwork']}")

    def toggle_dry_run(self):
        current = self.config["dry_run"]
        self.config["dry_run"] = not current
        CONFIG_PATH.write_text(json.dumps(self.config, indent=2))
        self._log(logging.INFO, f"Dry run set to {self.config['dry_run']}")

    def refresh_artwork(self):
        if self.current_track_info:
            self._enqueue_artwork(self.current_track_info)

    def force_reconnect(self):
        self.connect()

    def _enqueue_artwork(self, track):
        try:
            while not self.artwork_queue.empty():
                self.artwork_queue.get_nowait()
        except queue.Empty:
            pass
        self.artwork_queue.put(track)

    def _artwork_worker(self):
        while True:
            track = self.artwork_queue.get()
            if track is None:
                break
            try:
                self.cached_artwork_url = self.fetch_artwork_url(track["artist"], track["album"], track["name"])
            except Exception as exc:
                self._log(logging.DEBUG, f"Artwork worker error: {exc}")
            finally:
                self.artwork_queue.task_done()

    def _clear_presence(self):
        try:
            if self.config.get("dry_run"):
                self._log(logging.INFO, "Dry-run: clear presence")
                return
            self.rpc.clear()
        except Exception as exc:
            self._log(logging.DEBUG, f"Clear presence failed: {exc}")

    def _build_buttons(self, track):
        if not self.config.get("buttons_enabled", True):
            return None
        safe_values = {}
        for key, val in track.items():
            try:
                safe_values[key] = urllib.parse.quote_plus(str(val)) if val is not None else ""
            except Exception:
                safe_values[key] = ""
        buttons = []
        for button in self.config.get("buttons", [])[:2]:
            try:
                url = button.get("url_template", "").format(**safe_values)
                if url:
                    buttons.append({"label": button.get("label", "Open"), "url": url})
            except Exception:
                continue
        return buttons or None

    def _apply_privacy(self, track):
        if not track:
            return None
        if not self.config["privacy"].get("hide_metadata"):
            return track
        return {
            "id": track["id"],
            "name": "Listening in iTunes",
            "artist": "Private",
            "album": "Private",
            "position": track.get("position", 0),
            "duration": track.get("duration", 0),
        }

    def _update_presence(self, track):
        safe_track = self._apply_privacy(track) if track else None
        if not safe_track:
            self._clear_presence()
            return

        play_count = self.play_counts.get(track["id"], 0)
        play_text = f"{play_count} play{'s' if play_count != 1 else ''} this session"

        privacy_on = self.config["privacy"].get("hide_metadata")
        if privacy_on:
            fmt_details = "Listening to music"
            fmt_state = None
        else:
            fmt_details = self.config["presence_format"]["details"].format(**safe_track, play_text=play_text)
            fmt_state = self.config["presence_format"]["state"].format(**safe_track, play_text=play_text)

        position = safe_track.get("position", 0) or 0
        duration = safe_track.get("duration", 0) or 0
        time_remaining = max(duration - position, 0)
        end_timestamp = int(time.time() + time_remaining) if duration else None
        start_timestamp = int(time.time() - position) if duration else None

        payload = {
            "details": fmt_details[:128],
            "state": fmt_state[:128] if fmt_state else None,
            "large_image": FALLBACK_IMAGE if privacy_on else self.cached_artwork_url,
            "large_text": None if privacy_on else safe_track.get("album"),
            "start": start_timestamp,
            "end": end_timestamp,
            "buttons": None if privacy_on else self._build_buttons(safe_track),
        }

        try:
            if self.config.get("dry_run"):
                self._log(logging.INFO, f"Dry-run payload: {payload}")
                return
            self.rpc.update(**{k: v for k, v in payload.items() if v is not None})
        except Exception as exc:
            self._log(logging.WARNING, f"Presence update failed: {exc}")
            self.ensure_connected()

    def loop(self):
        pythoncom.CoInitialize()
        while self.running:
            try:
                if not self.rpc_enabled:
                    time.sleep(self.config["idle_interval"])
                    continue

                track = self.get_track_info()
                self.current_track_info = track
                if track:
                    self.menu_track_info = track
                    self.menu_track_seen_at = time.time()
                else:
                    if time.time() - self.menu_track_seen_at > 10:
                        self.menu_track_info = None

                if track:
                    if track["id"] != self.last_track_id:
                        self.cached_artwork_url = FALLBACK_IMAGE
                        self._enqueue_artwork(track)
                        self.last_track_id = track["id"]
                        self.track_started_at = time.time()
                        self.track_counted = False

                    if not self.track_counted and self.track_started_at:
                        time_listened = time.time() - self.track_started_at
                        threshold = min(
                            self.config["play_count_threshold_seconds"],
                            track["duration"] * self.config["play_count_threshold_fraction"],
                        )
                        if time_listened >= threshold:
                            self.play_counts[track["id"]] = self.play_counts.get(track["id"], 0) + 1
                            self.track_counted = True

                    self._update_presence(track)
                else:
                    self._clear_presence()
                    self.last_track_id = None
                    self.track_started_at = None
                    self.track_counted = False

                time.sleep(self.config["refresh_interval"])

            except Exception as e:
                self._log(logging.WARNING, f"Loop error: {e}")
                time.sleep(self.config["idle_interval"])


config = load_config()
configure_logging()
handler = RPCHandler(config)

def get_menu_items():
    """Generates the menu items dynamically every time you right-click."""
    
    items = []
    track = handler.menu_track_info if handler.rpc_enabled else None

    toggle_label = "Enable RPC" if not handler.rpc_enabled else "Disable RPC"
    
    def on_toggle(icon, item):
        handler.toggle_rpc()

    def on_refresh(icon, item):
        handler.refresh_artwork()

    def on_reconnect(icon, item):
        handler.force_reconnect()

    def on_privacy(icon, item):
        handler.toggle_privacy()

    def on_dry_run(icon, item):
        handler.toggle_dry_run()

    def on_network(icon, item):
        handler.toggle_network_artwork()

    def on_debug(icon, item):
        logging.info(
            {
                "rpc_enabled": handler.rpc_enabled,
                "dry_run": handler.config.get("dry_run"),
                "privacy": handler.config.get("privacy"),
                "current_track": handler.current_track_info,
                "cached_artwork": handler.cached_artwork_url,
            }
        )

    items.append(pystray.MenuItem(toggle_label, on_toggle))
    items.append(pystray.MenuItem("Refresh Artwork", on_refresh))
    items.append(pystray.MenuItem("Force Reconnect", on_reconnect))
    items.append(pystray.MenuItem("Toggle Privacy", on_privacy))
    items.append(pystray.MenuItem("Toggle Dry Run", on_dry_run))
    items.append(pystray.MenuItem("Toggle Network Artwork", on_network))
    items.append(pystray.MenuItem("Log Debug Info", on_debug))
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