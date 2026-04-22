import argparse
import asyncio
import keyboard
import sys
import re
import os
import json
import win32com.client
from typing import Optional, List, Tuple
import winreg
from datetime import datetime
from xml.sax.saxutils import escape

try:
    from winrt.windows.media.speechsynthesis import SpeechSynthesizer
    from winrt.windows.media.playback import MediaPlayer, MediaPlayerAudioCategory
    WINRT_AVAILABLE = True
except ImportError:
    SpeechSynthesizer = None
    MediaPlayer = None
    MediaPlayerAudioCategory = None
    WINRT_AVAILABLE = False

# --- CONFIGURATION ---

# Map symbols to (Display Text, Spoken Text)
TOKEN_MAP = {
    '\n': ("<NEW LINE>", "new line"),
    '...': ("<ELLIPSIS>", "ellipsis"),
    '…': ("<ELLIPSIS>", "ellipsis"),
    '.': ("<PERIOD>", "period"),
    ',': ("<COMMA>", "comma"),
    ':': ("<COLON>", "colon"),
    ';': ("<SEMICOLON>", "semicolon"),
    '!': ("<EXCLAMATION>", "exclamation mark"),
    '?': ("<QUESTION>", "question mark"),
    '"': ("<QUOTE>", "quote"),
    '“': ("<QUOTE>", "quote"),
    '”': ("<QUOTE>", "quote"),
    '-': ("<DASH>", "dash"),
    '—': ("<EM DASH>", "dash"),
}

# Global settings are stored separately from per-book progress.
SETTINGS_FILE = os.path.join(os.path.dirname(__file__), "readword_settings.json")

# Voices keywords to prioritize (highest priority first).
# This strongly prefers newer-sounding Microsoft voices when available and
# avoids the old David desktop voice unless it is the only option.
PREFERRED_VOICE_TAGS = [
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
DISFAVORED_VOICE_TAGS = ["Desktop", "David"]

WINRT_VOICE_PROFILES = {
    "microsoft mark": {"rate": "+8%", "pitch": "+8%", "volume": "100"},
    "microsoft zira": {"rate": "+6%", "pitch": "+6%", "volume": "100"},
    "microsoft david": {"rate": "+4%", "pitch": "+2%", "volume": "100"},
}
DEFAULT_WINRT_PROFILE = {"rate": "+6%", "pitch": "+6%", "volume": "100"}


class WindowsSpeechEngine:
    def __init__(self, preferred_voice: Optional[str] = None, prompt_for_voice: bool = False, quiet: bool = False):
        """Initialize the best available Windows speech engine with voice selection."""
        self.settings_file = SETTINGS_FILE
        self.preferred_voice = preferred_voice
        self.prompt_for_voice = prompt_for_voice
        self.quiet = quiet
        self.backend = "sapi"
        self.player = None
        self.current_stream = None

        if not self._init_winrt():
            self._init_sapi()

        self._select_voice(preferred_voice, prompt_for_voice)

    def _init_winrt(self) -> bool:
        if not WINRT_AVAILABLE:
            return False

        try:
            self.backend = "winrt"
            self.synthesizer = SpeechSynthesizer()
            self.player = MediaPlayer()
            self.player.audio_category = MediaPlayerAudioCategory.SPEECH
            self.player.real_time_playback = True
            self.player.auto_play = False
            self.voices = list(SpeechSynthesizer.all_voices)
            return bool(self.voices)
        except Exception:
            return False

    def _init_sapi(self):
        self.backend = "sapi"
        self._refresh_speech_tokens()
        self.speaker = win32com.client.Dispatch("SAPI.SpVoice")
        self.speaker.Rate = 0
        self.speaker.Volume = 100
        raw_voices = self.speaker.GetVoices()
        self.voices = [raw_voices.Item(i) for i in range(raw_voices.Count)]

    def _refresh_speech_tokens(self):
        """Registry hack to make sure all voices appear."""
        try:
            path = r"SOFTWARE\Microsoft\Speech\Voices\Tokens"
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, path, 0, winreg.KEY_READ) as key:
                winreg.EnumKey(key, 0)
        except Exception:
            pass # Fail silently, this is just a helper

    def list_voice_descriptions(self) -> List[str]:
        return [self._voice_name(voice) for voice in self.voices]

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

    def _score_voice(self, description: str) -> int:
        normalized = description.casefold()
        score = 0

        for priority, tag in enumerate(PREFERRED_VOICE_TAGS):
            if tag.casefold() in normalized:
                score += 100 - (priority * 10)

        for tag in DISFAVORED_VOICE_TAGS:
            if tag.casefold() in normalized:
                score -= 35

        if "english (united states)" in normalized:
            score += 20
        if normalized.startswith("microsoft "):
            score += 5

        return score

    def _voice_name(self, voice) -> str:
        if self.backend == "winrt":
            return voice.display_name
        return voice.GetDescription()

    def _current_voice_name(self) -> str:
        if self.backend == "winrt":
            voice = self.synthesizer.voice
            return voice.display_name if voice else ""

        try:
            return self.speaker.Voice.GetDescription()
        except Exception:
            return ""

    def _build_ssml(self, text: str) -> str:
        voice_name = self._current_voice_name()
        normalized_voice_name = voice_name.casefold()
        profile = WINRT_VOICE_PROFILES.get(normalized_voice_name, DEFAULT_WINRT_PROFILE)
        escaped_text = escape(text)
        escaped_voice_name = escape(voice_name)

        return (
            '<speak version="1.0" xmlns="http://www.w3.org/2001/10/synthesis" xml:lang="en-US">'
            f'<voice name="{escaped_voice_name}">'
            f'<prosody rate="{profile["rate"]}" pitch="{profile["pitch"]}" volume="{profile["volume"]}">'
            f"{escaped_text}"
            "</prosody>"
            "</voice>"
            "</speak>"
        )

    def _find_voice(self, query: Optional[str]):
        if not query:
            return None

        normalized_query = query.casefold().strip()

        for voice in self.voices:
            if self._voice_name(voice).casefold() == normalized_query:
                return voice

        for voice in self.voices:
            if normalized_query in self._voice_name(voice).casefold():
                return voice

        return None

    def _pick_best_voice(self):
        if not self.voices:
            return None

        return max(
            self.voices,
            key=lambda voice: (self._score_voice(self._voice_name(voice)), self._voice_name(voice)),
        )

    def _maybe_prompt_for_voice(self, selected_voice):
        available_voices = self.list_voice_descriptions()
        selected_name = self._voice_name(selected_voice) if selected_voice else None

        print("\n--- Available Voices ---")
        for i, description in enumerate(available_voices, start=1):
            marker = " (recommended)" if selected_name and description == selected_name else ""
            print(f"{i}. {description}{marker}")

        print("\nPress ENTER to keep the recommended voice, or type a number (1-N) to switch:")
        choice = input("> ").strip()

        if choice.isdigit():
            idx = int(choice) - 1
            if 0 <= idx < len(self.voices):
                return self.voices[idx]

        return selected_voice

    def _select_voice(self, preferred_voice: Optional[str], prompt_for_voice: bool):
        """Find the best available voice and optionally let the user override it."""
        settings = self._load_settings()
        selected_voice = self._find_voice(preferred_voice)
        if preferred_voice and not selected_voice:
            print(f'Warning: voice "{preferred_voice}" was not found. Falling back to the best available voice.')

        if not selected_voice:
            saved_voice = settings.get("voice")
            selected_voice = self._find_voice(saved_voice)

        if not selected_voice:
            selected_voice = self._pick_best_voice()

        should_prompt = prompt_for_voice or (
            not preferred_voice
            and not settings.get("voice")
            and len(self.voices) > 1
            and not self.quiet
        )

        if should_prompt and self.voices:
            selected_voice = self._maybe_prompt_for_voice(selected_voice)

        if selected_voice:
            description = self._voice_name(selected_voice)
            if self.backend == "winrt":
                self.synthesizer.voice = selected_voice
            else:
                self.speaker.Voice = selected_voice
            self._save_settings({"voice": description})
            if not self.quiet:
                backend_name = "WinRT" if self.backend == "winrt" else "SAPI"
                print(f"Selected voice [{backend_name}]: {description}")
        elif not self.quiet:
            print("Using system default voice.")

    def say(self, text):
        """Speak asynchronously."""
        if self.backend == "winrt":
            try:
                if self.player:
                    self.player.pause()
            except Exception:
                pass

            async def synthesize():
                ssml = self._build_ssml(text)
                return await self.synthesizer.synthesize_ssml_to_stream_async(ssml)

            self.current_stream = asyncio.run(synthesize())
            self.player.set_stream_source(self.current_stream)
            self.player.play()
            return

        self.speaker.Speak(text, 1) # 1 = SVSFlagsAsync

    def wait(self):
        """Block until speaking is done."""
        if self.backend == "winrt":
            return
        self.speaker.WaitUntilDone(10000) # Timeout after 10s to prevent hanging


class BookReader:
    def __init__(self, filepath, preferred_voice: Optional[str] = None, prompt_for_voice: bool = False):
        self.filepath = filepath
        self.save_file = os.path.splitext(filepath)[0] + "_progress.json"
        self.engine = WindowsSpeechEngine(
            preferred_voice=preferred_voice,
            prompt_for_voice=prompt_for_voice,
        )
        self.tokens = self._tokenize_file()
        self.position = self._load_progress()
        self.is_running = True

    def _tokenize_file(self) -> List[str]:
        """Reads file and returns a list of actionable tokens."""
        if not os.path.exists(self.filepath):
            print(f"Error: File {self.filepath} not found.")
            sys.exit(1)

        with open(self.filepath, 'r', encoding='utf-8') as f:
            text = f.read()

        # Regex Explanation:
        # 1. \.\.\.           -> Ellipsis (...)
        # 2. [\w'’]+          -> Words, including STRAIGHT (') and CURLY (’) apostrophes
        # 3. [^\w\s]          -> Any single character that isn't a word or whitespace
        # 4. \n               -> Newlines (explicitly captured)
        pattern = r"\.\.\.|[\w'’]+|[^\w\s]|\n"
        return [t for t in re.findall(pattern, text) if t.strip() or t == '\n']

    def _load_progress(self) -> int:
        """Loads index from JSON."""
        if os.path.exists(self.save_file):
            try:
                with open(self.save_file, 'r') as f:
                    data = json.load(f)
                    pos = data.get('position', 0)
                    return max(0, min(pos, len(self.tokens) - 1))
            except Exception:
                pass
        return 0

    def _save_progress(self):
        """Saves current index to JSON."""
        try:
            with open(self.save_file, 'w') as f:
                json.dump({
                    'position': self.position, 
                    'timestamp': str(datetime.now())
                }, f)
        except Exception as e:
            print(f"Failed to save progress: {e}")

    def _get_display_info(self, token) -> Tuple[str, str]:
        """Returns (Display Text, Spoken Text) for a token."""
        # Check dictionary first
        if token in TOKEN_MAP:
            return TOKEN_MAP[token]
        
        # Check for curly quotes often found in books
        if token in ['“', '”', '"']: 
            return ("<QUOTE>", "quote")
            
        # Default: It's a word
        return (token, token)

    def run(self):
        print("\n" + "="*40)
        print(f"Loaded {len(self.tokens)} tokens.")
        print("CONTROLS: [SPACE] Next  |  [BACKSPACE] Previous  |  [ESC] Quit")
        print("="*40 + "\n")

        # Show initial context
        display, _ = self._get_display_info(self.tokens[self.position])
        print(f"Ready at {self.position}/{len(self.tokens)}: {display}")

        while self.is_running:
            # wait_until blocks nicely, saving CPU
            event = keyboard.read_event(suppress=True)

            if event.event_type == keyboard.KEY_DOWN:
                if event.name == 'esc':
                    print("\nExiting...")
                    self._save_progress()
                    self.is_running = False
                
                elif event.name == 'space':
                    if self.position < len(self.tokens):
                        token = self.tokens[self.position]
                        disp_text, speak_text = self._get_display_info(token)
                        
                        print(f"[{self.position}] {disp_text}")
                        self.engine.say(speak_text)
                        
                        self.position += 1
                        
                        # Saving on every single step (No optimization)
                        self._save_progress()
                    else:
                        print("End of file reached.")

                elif event.name == 'backspace':
                    if self.position > 0:
                        self.position -= 1
                        token = self.tokens[self.position]
                        disp_text, _ = self._get_display_info(token)
                        print(f"<- Back to: {disp_text}")
                        # Saving on every single step
                        self._save_progress()

def parse_args():
    parser = argparse.ArgumentParser(description="Read a text file one token at a time.")
    parser.add_argument("file", nargs="?", help="Path to the text file to read.")
    parser.add_argument(
        "--voice",
        help="Voice name or partial description to use. Example: --voice Zira",
    )
    parser.add_argument(
        "--choose-voice",
        action="store_true",
        help="Show the available voices at startup and let you override the recommendation.",
    )
    parser.add_argument(
        "--list-voices",
        action="store_true",
        help="List detected voices and exit.",
    )
    return parser.parse_args()


if __name__ == "__main__":
    args = parse_args()

    if args.list_voices:
        engine = WindowsSpeechEngine(preferred_voice=args.voice, prompt_for_voice=False, quiet=True)
        print("\nDetected voices:")
        for description in engine.list_voice_descriptions():
            print(f"- {description}")
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
        )
        app.run()
    except KeyboardInterrupt:
        print("\nInterrupted by user.")
    except Exception as e:
        print(f"\nCritical Error: {e}")
