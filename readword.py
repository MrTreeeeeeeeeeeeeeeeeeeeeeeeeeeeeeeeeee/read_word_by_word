import keyboard
import sys
import re
import os
import json
import win32com.client
from typing import Optional, List, Dict, Tuple
import winreg
from datetime import datetime

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

# Voices keywords to prioritize (in order)
PREFERRED_VOICE_TAGS = ["Natural", "Neural", "Mobile", "Microsoft"]


class WindowsSpeechEngine:
    def __init__(self):
        """Initialize Windows SAPI speech engine with voice selection."""
        self._refresh_speech_tokens()
        self.speaker = win32com.client.Dispatch("SAPI.SpVoice")
        self.speaker.Rate = 0
        self.speaker.Volume = 100
        
        self.voices = self.speaker.GetVoices()
        self._select_best_voice()

    def _refresh_speech_tokens(self):
        """Registry hack to make sure all voices appear."""
        try:
            path = r"SOFTWARE\Microsoft\Speech\Voices\Tokens"
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, path, 0, winreg.KEY_READ) as key:
                winreg.EnumKey(key, 0)
        except Exception:
            pass # Fail silently, this is just a helper

    def _select_best_voice(self):
        """Finds the best available voice or prompts user."""
        available_voices = []
        print("\n--- Available Voices ---")
        
        for i, voice in enumerate(self.voices):
            desc = voice.GetDescription()
            available_voices.append(voice)
            print(f"{i + 1}. {desc}")

        # Auto-select based on priority list
        selected_voice = None
        for tag in PREFERRED_VOICE_TAGS:
            for voice in available_voices:
                if tag.lower() in voice.GetDescription().lower():
                    selected_voice = voice
                    break
            if selected_voice: break

        # If we found a good auto-candidate, mention it
        if selected_voice:
             print(f"\nRecommended: {selected_voice.GetDescription()}")

        print("\nPress ENTER to use recommended, or type a number (1-N) to switch:")
        choice = input("> ").strip()
        
        if choice.isdigit():
            idx = int(choice) - 1
            if 0 <= idx < len(available_voices):
                selected_voice = available_voices[idx]

        # Fallback to system default if nothing selected/found
        if selected_voice:
            self.speaker.Voice = selected_voice
            print(f"Selected: {selected_voice.GetDescription()}")
        else:
            print("Using system default voice.")

    def say(self, text):
        """Speak asynchronously."""
        self.speaker.Speak(text, 1) # 1 = SVSFlagsAsync

    def wait(self):
        """Block until speaking is done."""
        self.speaker.WaitUntilDone(10000) # Timeout after 10s to prevent hanging


class BookReader:
    def __init__(self, filepath):
        self.filepath = filepath
        self.save_file = os.path.splitext(filepath)[0] + "_progress.json"
        self.engine = WindowsSpeechEngine()
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

if __name__ == "__main__":
    target_file = ""
    
    # Argument parsing
    if len(sys.argv) > 1:
        potential_path = sys.argv[1]
        if os.path.exists(potential_path):
            target_file = potential_path
    
    if not target_file:
        target_file = input("Enter path to text file: ").strip().strip('"')

    try:
        app = BookReader(target_file)
        app.run()
    except KeyboardInterrupt:
        print("\nInterrupted by user.")
    except Exception as e:
        print(f"\nCritical Error: {e}")