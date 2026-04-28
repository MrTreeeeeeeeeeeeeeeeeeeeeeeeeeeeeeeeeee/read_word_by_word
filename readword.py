import argparse
import asyncio
import concurrent.futures
import ctypes
import hashlib
import json
import os
import sys
import threading
from datetime import datetime
from typing import Dict, List, Optional

import keyboard
import win32com.client
import winreg

from readword_core import EOF_MESSAGE, ReaderSession

try:
    import edge_tts

    EDGE_TTS_AVAILABLE = True
except ImportError:
    edge_tts = None
    EDGE_TTS_AVAILABLE = False

try:
    from winrt.windows.media.playback import MediaPlayer, MediaPlayerAudioCategory
    from winrt.windows.media.speechsynthesis import SpeechSynthesizer

    WINRT_AVAILABLE = True
except ImportError:
    MediaPlayer = None
    MediaPlayerAudioCategory = None
    SpeechSynthesizer = None
    WINRT_AVAILABLE = False


# --- CONFIGURATION ---

APP_DIR = os.path.dirname(__file__)
SETTINGS_FILE = os.path.join(APP_DIR, "readword_settings.json")
CACHE_DIR = os.path.join(APP_DIR, ".tts_cache")
EDGE_PLAYBACK_ALIAS = "readword_tts"
EDGE_STALE_TEMP_SECONDS = 120

PREFERRED_EDGE_VOICES = [
    "en-US-EmmaMultilingualNeural",
    "en-US-AndrewNeural",
    "en-US-AvaNeural",
    "en-US-BrianMultilingualNeural",
    "en-US-AnaNeural",
    "en-US-JennyNeural",
]
EDGE_VOICE_ALIASES = {
    "female": "en-US-EmmaMultilingualNeural",
    "woman": "en-US-EmmaMultilingualNeural",
    "male": "en-US-AndrewNeural",
    "man": "en-US-AndrewNeural",
}

PREFERRED_LOCAL_VOICE_TAGS = [
    "Aria",
    "Jenny",
    "Guy",
    "Sonia",
    "Mark",
    "Zira",
    "Natural",
    "Neural",
    "Mobile",
]
DISFAVORED_LOCAL_VOICE_TAGS = ["Desktop", "David"]


class WindowsSpeechEngine:
    def __init__(
        self,
        preferred_voice: Optional[str] = None,
        prompt_for_voice: bool = False,
        quiet: bool = False,
        backend_preference: str = "auto",
    ):
        self.settings_file = SETTINGS_FILE
        self.quiet = quiet
        self.backend = "sapi"
        self.backend_preference = backend_preference
        self.current_stream = None
        self.player = None
        self.edge_voice_catalog: Optional[List[dict]] = None
        self.edge_executor: Optional[concurrent.futures.ThreadPoolExecutor] = None
        self.edge_futures: Dict[str, concurrent.futures.Future] = {}
        self.edge_lock = threading.Lock()
        self.winmm = ctypes.windll.winmm

        os.makedirs(CACHE_DIR, exist_ok=True)

        if not self._initialize_backend(backend_preference):
            raise RuntimeError("No speech backend could be initialized.")

        self._repair_cache(startup=True)
        self._select_voice(preferred_voice, prompt_for_voice)

    def _initialize_backend(self, backend_preference: str) -> bool:
        backend_order = []
        if backend_preference == "edge":
            backend_order = ["edge", "winrt", "sapi"]
        elif backend_preference == "winrt":
            backend_order = ["winrt", "sapi"]
        elif backend_preference == "sapi":
            backend_order = ["sapi"]
        else:
            backend_order = ["edge", "winrt", "sapi"]

        for backend_name in backend_order:
            try:
                if backend_name == "edge" and self._init_edge():
                    return True
                if backend_name == "winrt" and self._init_winrt():
                    return True
                if backend_name == "sapi" and self._init_sapi():
                    return True
            except Exception:
                continue

        return False

    def _init_edge(self) -> bool:
        if not EDGE_TTS_AVAILABLE:
            return False

        self.backend = "edge"
        self.selected_voice = None
        self.edge_executor = concurrent.futures.ThreadPoolExecutor(
            max_workers=3,
            thread_name_prefix="readword_edge_tts",
        )
        self.voices = list(PREFERRED_EDGE_VOICES)
        return True

    def _init_winrt(self) -> bool:
        if not WINRT_AVAILABLE:
            return False

        self.backend = "winrt"
        self.synthesizer = SpeechSynthesizer()
        self.player = MediaPlayer()
        self.player.audio_category = MediaPlayerAudioCategory.SPEECH
        self.player.real_time_playback = True
        self.player.auto_play = False
        self.voices = list(SpeechSynthesizer.all_voices)
        return bool(self.voices)

    def _init_sapi(self) -> bool:
        self.backend = "sapi"
        self._refresh_speech_tokens()
        self.speaker = win32com.client.Dispatch("SAPI.SpVoice")
        self.speaker.Rate = 0
        self.speaker.Volume = 100
        raw_voices = self.speaker.GetVoices()
        self.voices = [raw_voices.Item(i) for i in range(raw_voices.Count)]
        return bool(self.voices)

    def _refresh_speech_tokens(self):
        try:
            path = r"SOFTWARE\Microsoft\Speech\Voices\Tokens"
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, path, 0, winreg.KEY_READ) as key:
                winreg.EnumKey(key, 0)
        except Exception:
            pass

    def _load_settings(self) -> dict:
        if not os.path.exists(self.settings_file):
            return {}

        try:
            with open(self.settings_file, "r", encoding="utf-8") as f:
                data = json.load(f)
                return data if isinstance(data, dict) else {}
        except Exception:
            return {}

    def _save_settings(self, updates: dict):
        settings = self._load_settings()
        settings.update(updates)

        try:
            with open(self.settings_file, "w", encoding="utf-8") as f:
                json.dump(settings, f, indent=2)
        except Exception:
            pass

    def _load_edge_voice_catalog(self) -> List[dict]:
        if self.edge_voice_catalog is not None:
            return self.edge_voice_catalog

        if self.backend != "edge":
            return []

        async def load():
            return await edge_tts.list_voices()

        voices = asyncio.run(load())
        self.edge_voice_catalog = [voice for voice in voices if voice.get("Locale") == "en-US"]
        return self.edge_voice_catalog

    def _voice_name(self, voice) -> str:
        if self.backend == "edge":
            return voice
        if self.backend == "winrt":
            return voice.display_name
        return voice.GetDescription()

    def _edge_voice_entry(self, short_name: str) -> Optional[dict]:
        if self.backend != "edge":
            return None

        for entry in self._load_edge_voice_catalog():
            if entry.get("ShortName") == short_name:
                return entry

        return None

    def _voice_label(self, voice) -> str:
        if self.backend != "edge":
            return self._voice_name(voice)

        short_name = self._voice_name(voice)
        entry = self._edge_voice_entry(short_name)
        if entry:
            gender = entry.get("Gender")
            gender_label = f" [{gender}]" if gender else ""
            return f"{short_name} - {entry.get('FriendlyName')}{gender_label}"

        return short_name

    def list_voice_descriptions(self) -> List[str]:
        if self.backend == "edge":
            return [self._voice_label(voice) for voice in self._load_edge_voice_catalog_entries()]

        return [self._voice_name(voice) for voice in self.voices]

    def current_voice_name(self) -> str:
        if self.backend == "edge":
            return self.selected_voice or ""
        if self.backend == "winrt":
            return self.synthesizer.voice.display_name if self.synthesizer.voice else ""
        return self.speaker.Voice.GetDescription() if self.speaker.Voice else ""

    def current_voice_label(self) -> str:
        if self.backend == "edge":
            return self._voice_label(self.selected_voice) if self.selected_voice else ""
        if self.backend == "winrt":
            return self._voice_label(self.synthesizer.voice) if self.synthesizer.voice else ""
        return self._voice_label(self.speaker.Voice) if self.speaker.Voice else ""

    def _load_edge_voice_catalog_entries(self) -> List[str]:
        catalog = self._load_edge_voice_catalog()
        ordered = []
        seen = set()

        for short_name in PREFERRED_EDGE_VOICES:
            for entry in catalog:
                if entry.get("ShortName") == short_name and short_name not in seen:
                    ordered.append(short_name)
                    seen.add(short_name)

        for entry in catalog:
            short_name = entry.get("ShortName")
            if short_name and short_name not in seen:
                ordered.append(short_name)
                seen.add(short_name)

        return ordered

    def _score_local_voice(self, description: str) -> int:
        normalized = description.casefold()
        score = 0

        for priority, tag in enumerate(PREFERRED_LOCAL_VOICE_TAGS):
            if tag.casefold() in normalized:
                score += 100 - (priority * 10)

        for tag in DISFAVORED_LOCAL_VOICE_TAGS:
            if tag.casefold() in normalized:
                score -= 35

        if "english (united states)" in normalized:
            score += 20

        return score

    def _find_voice(self, query: Optional[str]):
        if not query:
            return None

        normalized_query = query.casefold().strip()

        if self.backend == "edge":
            alias_target = EDGE_VOICE_ALIASES.get(normalized_query)
            if alias_target:
                normalized_query = alias_target.casefold()

            for voice in self._load_edge_voice_catalog_entries():
                if voice.casefold() == normalized_query:
                    return voice

            for entry in self._load_edge_voice_catalog():
                friendly = entry.get("FriendlyName", "").casefold()
                short_name = entry.get("ShortName", "")
                if normalized_query in short_name.casefold() or normalized_query in friendly:
                    return short_name

            return None

        for voice in self.voices:
            description = self._voice_name(voice).casefold()
            if description == normalized_query or normalized_query in description:
                return voice

        return None

    def _pick_best_voice(self):
        if self.backend == "edge":
            catalog_voices = set(self._load_edge_voice_catalog_entries())
            for voice in PREFERRED_EDGE_VOICES:
                if voice in catalog_voices:
                    return voice
            return self._load_edge_voice_catalog_entries()[0] if catalog_voices else PREFERRED_EDGE_VOICES[0]

        if not self.voices:
            return None

        return max(
            self.voices,
            key=lambda voice: (self._score_local_voice(self._voice_name(voice)), self._voice_name(voice)),
        )

    def _prompt_voice_choices(self, selected_voice) -> List:
        if self.backend != "edge":
            return self.voices

        available = self._load_edge_voice_catalog_entries()
        selected_name = self._voice_name(selected_voice) if selected_voice else None
        choices = []
        seen = set()

        explicit_priority = [
            EDGE_VOICE_ALIASES["female"],
            EDGE_VOICE_ALIASES["male"],
        ]

        for voice in explicit_priority + PREFERRED_EDGE_VOICES:
            if voice in available and voice not in seen:
                choices.append(voice)
                seen.add(voice)

        if selected_name and selected_name not in seen:
            choices.append(selected_name)

        return choices

    def _maybe_prompt_for_voice(self, selected_voice):
        choices = self._prompt_voice_choices(selected_voice)
        selected_name = self._voice_name(selected_voice) if selected_voice else None

        if not choices:
            return selected_voice

        print("\n--- Available Voices ---")
        for i, voice in enumerate(choices, start=1):
            label = self._voice_label(voice)
            marker = " (recommended)" if selected_name and self._voice_name(voice) == selected_name else ""
            if self._voice_name(voice) == EDGE_VOICE_ALIASES["female"]:
                marker += " (female)"
            elif self._voice_name(voice) == EDGE_VOICE_ALIASES["male"]:
                marker += " (male)"
            print(f"{i}. {label}{marker}")

        print("\nPress ENTER to keep the recommended voice, or type a number (1-N) to switch:")
        choice = input("> ").strip()

        if choice.isdigit():
            idx = int(choice) - 1
            if 0 <= idx < len(choices):
                return choices[idx]

        return selected_voice

    def _select_voice(self, preferred_voice: Optional[str], prompt_for_voice: bool):
        settings = self._load_settings()
        selected_voice = self._find_voice(preferred_voice)

        if preferred_voice and not selected_voice and not self.quiet:
            print(f'Warning: voice "{preferred_voice}" was not found. Falling back to the best available voice.')

        if not selected_voice:
            saved_voice = settings.get("voice")
            selected_voice = self._find_voice(saved_voice)

        if not selected_voice:
            selected_voice = self._pick_best_voice()

        should_prompt = (
            not preferred_voice
            and len(self._prompt_voice_choices(selected_voice)) > 1
            and not self.quiet
        ) or prompt_for_voice

        if should_prompt:
            selected_voice = self._maybe_prompt_for_voice(selected_voice)

        if not selected_voice:
            raise RuntimeError("No usable voice was found for the selected backend.")

        if self.backend == "edge":
            self.selected_voice = selected_voice
        elif self.backend == "winrt":
            self.synthesizer.voice = selected_voice
        else:
            self.speaker.Voice = selected_voice

        self._save_settings({"voice": self._voice_name(selected_voice), "backend": self.backend})

        if not self.quiet:
            print(f"Selected voice [{self.backend.upper()}]: {self._voice_label(selected_voice)}")

    def _edge_cache_key(self, text: str) -> str:
        payload = f"{self.selected_voice}|{text}".encode("utf-8")
        return hashlib.sha1(payload).hexdigest()

    def _edge_cache_path(self, text: str) -> str:
        key = self._edge_cache_key(text)
        filename = f"{key}.mp3"
        return os.path.join(CACHE_DIR, filename)

    def _is_valid_edge_cache_file(self, path: str) -> bool:
        try:
            return os.path.isfile(path) and os.path.getsize(path) > 0
        except OSError:
            return False

    def _cleanup_edge_temp_files(self, stale_only: bool):
        if not os.path.isdir(CACHE_DIR):
            return

        now = datetime.now().timestamp()
        for name in os.listdir(CACHE_DIR):
            path = os.path.join(CACHE_DIR, name)
            if not os.path.isfile(path):
                continue

            is_temp = ".tmp-" in name
            is_empty_mp3 = name.endswith(".mp3") and not self._is_valid_edge_cache_file(path)
            if not is_temp and not is_empty_mp3:
                continue

            if stale_only:
                try:
                    age_seconds = now - os.path.getmtime(path)
                except OSError:
                    age_seconds = EDGE_STALE_TEMP_SECONDS + 1
                if age_seconds < EDGE_STALE_TEMP_SECONDS:
                    continue

            try:
                os.remove(path)
            except OSError:
                pass

    def _repair_cache(self, startup: bool = False):
        if self.backend != "edge":
            return

        self._cleanup_edge_temp_files(stale_only=startup)

    def _synthesize_edge_file(self, text: str) -> str:
        output_path = self._edge_cache_path(text)
        if self._is_valid_edge_cache_file(output_path):
            return output_path

        if os.path.exists(output_path):
            try:
                os.remove(output_path)
            except OSError:
                pass

        temp_path = f"{output_path}.tmp-{threading.get_ident()}"
        communicate = edge_tts.Communicate(
            text=text,
            voice=self.selected_voice,
            rate="+8%",
            volume="+0%",
            pitch="+0Hz",
        )
        try:
            communicate.save_sync(temp_path)
            if not self._is_valid_edge_cache_file(temp_path):
                raise RuntimeError("Generated audio file was empty.")
            os.replace(temp_path, output_path)
        finally:
            if os.path.exists(temp_path):
                try:
                    os.remove(temp_path)
                except OSError:
                    pass
        return output_path

    def _submit_edge_prefetch(self, text: str):
        if self.backend != "edge" or not text:
            return

        cache_path = self._edge_cache_path(text)
        if self._is_valid_edge_cache_file(cache_path):
            return

        if os.path.exists(cache_path):
            try:
                os.remove(cache_path)
            except OSError:
                pass

        key = self._edge_cache_key(text)
        with self.edge_lock:
            future = self.edge_futures.get(key)
            if future and not future.done():
                return

            self.edge_futures[key] = self.edge_executor.submit(self._synthesize_edge_file, text)

    def _ensure_edge_audio(self, text: str) -> str:
        cache_path = self._edge_cache_path(text)
        if self._is_valid_edge_cache_file(cache_path):
            return cache_path

        if os.path.exists(cache_path):
            try:
                os.remove(cache_path)
            except OSError:
                pass

        key = self._edge_cache_key(text)
        with self.edge_lock:
            future = self.edge_futures.get(key)
            if future is None:
                future = self.edge_executor.submit(self._synthesize_edge_file, text)
                self.edge_futures[key] = future

        try:
            return future.result(timeout=20)
        finally:
            with self.edge_lock:
                current = self.edge_futures.get(key)
                if current is future and future.done():
                    self.edge_futures.pop(key, None)

    def _mci_error_message(self, code: int) -> str:
        buffer = ctypes.create_unicode_buffer(255)
        if self.winmm.mciGetErrorStringW(code, buffer, len(buffer)):
            return buffer.value
        return f"MCI error {code}"

    def _mci_command(self, command: str):
        buffer = ctypes.create_unicode_buffer(255)
        error_code = self.winmm.mciSendStringW(command, buffer, len(buffer), 0)
        if error_code:
            raise RuntimeError(self._mci_error_message(error_code))
        return buffer.value

    def _stop_edge_playback(self):
        try:
            self._mci_command(f"close {EDGE_PLAYBACK_ALIAS}")
        except Exception:
            pass

    def _play_edge_file(self, path: str):
        self._stop_edge_playback()
        normalized_path = os.path.abspath(path)
        self._mci_command(f'open "{normalized_path}" alias {EDGE_PLAYBACK_ALIAS}')
        self._mci_command(f"play {EDGE_PLAYBACK_ALIAS}")

    def prefetch(self, texts: List[str]):
        if self.backend != "edge":
            return

        seen = set()
        for text in texts:
            if not text or text in seen:
                continue
            seen.add(text)
            self._submit_edge_prefetch(text)

    def say(self, text: str):
        if self.backend == "edge":
            path = self._ensure_edge_audio(text)
            self._play_edge_file(path)
            return

        if self.backend == "winrt":
            try:
                if self.player:
                    self.player.pause()
            except Exception:
                pass

            async def synthesize():
                return await self.synthesizer.synthesize_text_to_stream_async(text)

            self.current_stream = asyncio.run(synthesize())
            self.player.set_stream_source(self.current_stream)
            self.player.play()
            return

        self.speaker.Speak(text, 1)

    def shutdown(self):
        if self.backend == "edge":
            self._stop_edge_playback()
            if self.edge_executor:
                self.edge_executor.shutdown(wait=False, cancel_futures=True)
            self._repair_cache(startup=False)
            return

        if self.backend == "winrt":
            try:
                if self.player:
                    self.player.pause()
                    self.player.close()
            except Exception:
                pass


class BookReader:
    def __init__(
        self,
        filepath: str,
        preferred_voice: Optional[str] = None,
        prompt_for_voice: bool = False,
        backend_preference: str = "auto",
    ):
        self.filepath = filepath
        self.engine = WindowsSpeechEngine(
            preferred_voice=preferred_voice,
            prompt_for_voice=prompt_for_voice,
            backend_preference=backend_preference,
        )
        self.session = ReaderSession(
            filepath,
            on_changed_file=self._should_reset_progress_for_changed_text,
        )
        self.is_running = True
        self._prime_audio_cache()

    def _should_reset_progress_for_changed_text(self) -> bool:
        print("\nThe text file changed since the last saved progress.")
        print("Reset saved progress and start from the beginning? [Y/n]")
        choice = input("> ").strip().lower()
        return choice in {"", "y", "yes"}

    def _save_progress(self):
        try:
            self.session.save_progress()
        except Exception as e:
            print(f"Failed to save progress: {e}")

    def _prime_audio_cache(self):
        self.engine.prefetch(self.session.prefetch_texts())

    def run(self):
        print("\n" + "=" * 40)
        print(f"Loaded {len(self.session.tokens)} tokens.")
        print("CONTROLS: [SPACE] Next  |  [BACKSPACE] Previous  |  [ESC] Quit")
        print("=" * 40 + "\n")

        current_token = self.session.current()
        print(f"Ready at {self.session.position}/{len(self.session.tokens)}: {current_token.display_text}")

        while self.is_running:
            event = keyboard.read_event(suppress=True)

            if event.event_type != keyboard.KEY_DOWN:
                continue

            if event.name == "esc":
                print("\nExiting...")
                self._save_progress()
                self.engine.shutdown()
                self.is_running = False
                continue

            if event.name == "space":
                token = self.session.next()
                if token.is_eof:
                    print(EOF_MESSAGE)
                    self.engine.say(EOF_MESSAGE)
                    continue

                print(f"[{token.position}] {token.display_text}")
                self.engine.say(token.speak_text)
                self._prime_audio_cache()
                continue

            if event.name == "backspace" and self.session.position > 0:
                token = self.session.previous()
                print(f"<- Back to: {token.display_text}")
                self._prime_audio_cache()


def parse_args():
    parser = argparse.ArgumentParser(description="Read a text file one token at a time.")
    parser.add_argument("file", nargs="?", help="Path to the text file to read.")
    parser.add_argument(
        "--voice",
        help="Voice name or partial description to use. Example: --voice Ava, --voice female, or --voice male",
    )
    parser.add_argument(
        "--choose-voice",
        action="store_true",
        help="Show the available voices at startup and let you override the recommendation.",
    )
    parser.add_argument(
        "--list-voices",
        action="store_true",
        help="List detected voices for the selected backend and exit.",
    )
    parser.add_argument(
        "--backend",
        choices=["auto", "edge", "winrt", "sapi"],
        default="auto",
        help="Speech backend to use. 'auto' prefers edge neural voices, then local Windows voices.",
    )
    return parser.parse_args()


if __name__ == "__main__":
    args = parse_args()

    if args.list_voices:
        engine = WindowsSpeechEngine(
            preferred_voice=args.voice,
            prompt_for_voice=False,
            quiet=True,
            backend_preference=args.backend,
        )
        print(f"\nDetected voices [{engine.backend.upper()}]:")
        for description in engine.list_voice_descriptions():
            print(f"- {description}")
        engine.shutdown()
        sys.exit(0)

    target_file = args.file or ""

    if target_file and not os.path.exists(target_file):
        print(f"Error: File {target_file} not found.")
        sys.exit(1)

    if not target_file:
        target_file = input("Enter path to text file: ").strip().strip('"')

    try:
        app = BookReader(
            target_file,
            preferred_voice=args.voice,
            prompt_for_voice=args.choose_voice,
            backend_preference=args.backend,
        )
        app.run()
    except KeyboardInterrupt:
        print("\nInterrupted by user.")
    except Exception as e:
        print(f"\nCritical Error: {e}")
