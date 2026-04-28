import concurrent.futures
import os
import sys
from tkinter import filedialog, messagebox

import customtkinter as ctk

from readword import WindowsSpeechEngine
from readword_core import EOF_MESSAGE, ReaderSession


BACKENDS = ["auto", "edge", "winrt", "sapi"]


class ReadWordApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        ctk.set_appearance_mode("system")
        ctk.set_default_color_theme("blue")

        self.title("Read Word by Word")
        self.geometry("900x520")
        self.minsize(720, 460)

        self.session = None
        self.engine = None
        self.displayed_token = None
        self.speech_executor = concurrent.futures.ThreadPoolExecutor(
            max_workers=1,
            thread_name_prefix="readword_ui_speech",
        )
        self.updating_voice = False

        self.file_var = ctk.StringVar(value="No file loaded")
        self.backend_var = ctk.StringVar(value="auto")
        self.voice_var = ctk.StringVar(value="")
        self.status_var = ctk.StringVar(value="Open a text file to begin.")
        self.progress_var = ctk.StringVar(value="0 / 0")

        self._build_layout()
        self._bind_keys()
        self._set_reader_controls_enabled(False)

        self.protocol("WM_DELETE_WINDOW", self.close)

    def _build_layout(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        top_bar = ctk.CTkFrame(self, corner_radius=0)
        top_bar.grid(row=0, column=0, sticky="ew")
        top_bar.grid_columnconfigure(1, weight=1)

        self.open_button = ctk.CTkButton(top_bar, text="Open File", command=self.open_file, width=110)
        self.open_button.grid(row=0, column=0, padx=(16, 10), pady=14)

        self.file_label = ctk.CTkLabel(top_bar, textvariable=self.file_var, anchor="w")
        self.file_label.grid(row=0, column=1, sticky="ew", padx=(0, 12))

        self.backend_menu = ctk.CTkOptionMenu(
            top_bar,
            values=BACKENDS,
            variable=self.backend_var,
            command=self._backend_changed,
            width=110,
        )
        self.backend_menu.grid(row=0, column=2, padx=(0, 10), pady=14)

        self.voice_menu = ctk.CTkOptionMenu(
            top_bar,
            values=[""],
            variable=self.voice_var,
            command=self._voice_changed,
            width=280,
        )
        self.voice_menu.grid(row=0, column=3, padx=(0, 16), pady=14)

        content = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        content.grid(row=1, column=0, sticky="nsew", padx=22, pady=(22, 10))
        content.grid_columnconfigure(0, weight=1)
        content.grid_rowconfigure(0, weight=1)

        token_panel = ctk.CTkFrame(content, corner_radius=8)
        token_panel.grid(row=0, column=0, sticky="nsew")
        token_panel.grid_columnconfigure(0, weight=1)
        token_panel.grid_rowconfigure(0, weight=1)

        self.token_label = ctk.CTkLabel(
            token_panel,
            text="Open a text file",
            font=ctk.CTkFont(size=46, weight="bold"),
            anchor="center",
            justify="center",
            wraplength=760,
        )
        self.token_label.grid(row=0, column=0, sticky="nsew", padx=28, pady=(28, 8))

        self.speak_preview_label = ctk.CTkLabel(
            token_panel,
            text="",
            font=ctk.CTkFont(size=18),
            text_color=("gray35", "gray75"),
            anchor="center",
            justify="center",
            wraplength=760,
        )
        self.speak_preview_label.grid(row=1, column=0, sticky="ew", padx=28, pady=(0, 24))

        progress_row = ctk.CTkFrame(content, corner_radius=0, fg_color="transparent")
        progress_row.grid(row=1, column=0, sticky="ew", pady=(18, 8))
        progress_row.grid_columnconfigure(0, weight=1)

        self.progress_bar = ctk.CTkProgressBar(progress_row)
        self.progress_bar.grid(row=0, column=0, sticky="ew", padx=(0, 12))
        self.progress_bar.set(0)

        self.progress_label = ctk.CTkLabel(progress_row, textvariable=self.progress_var, width=96)
        self.progress_label.grid(row=0, column=1)

        controls = ctk.CTkFrame(content, corner_radius=0, fg_color="transparent")
        controls.grid(row=2, column=0, sticky="ew", pady=(8, 0))
        controls.grid_columnconfigure((0, 1, 2, 3, 4), weight=1, uniform="controls")

        self.previous_button = ctk.CTkButton(controls, text="Previous", command=self.previous_token)
        self.previous_button.grid(row=0, column=0, sticky="ew", padx=(0, 8))

        self.speak_next_button = ctk.CTkButton(controls, text="Speak / Next", command=self.speak_next)
        self.speak_next_button.grid(row=0, column=1, sticky="ew", padx=8)

        self.repeat_button = ctk.CTkButton(controls, text="Repeat", command=self.repeat_token)
        self.repeat_button.grid(row=0, column=2, sticky="ew", padx=8)

        self.reset_button = ctk.CTkButton(controls, text="Reset", command=self.reset_progress)
        self.reset_button.grid(row=0, column=3, sticky="ew", padx=8)

        self.quit_button = ctk.CTkButton(controls, text="Quit", command=self.close)
        self.quit_button.grid(row=0, column=4, sticky="ew", padx=(8, 0))

        status_bar = ctk.CTkFrame(self, corner_radius=0)
        status_bar.grid(row=2, column=0, sticky="ew")
        status_bar.grid_columnconfigure(0, weight=1)

        self.status_label = ctk.CTkLabel(status_bar, textvariable=self.status_var, anchor="w")
        self.status_label.grid(row=0, column=0, sticky="ew", padx=16, pady=8)

        self.bind("<Configure>", self._resize_labels)

    def _bind_keys(self):
        self.bind("<space>", lambda event: self.speak_next())
        self.bind("<BackSpace>", lambda event: self.previous_token())
        self.bind("<r>", lambda event: self.repeat_token())
        self.bind("<R>", lambda event: self.repeat_token())
        self.bind("<Escape>", lambda event: self.close())

    def _resize_labels(self, event):
        if event.widget is not self:
            return
        wraplength = max(360, self.winfo_width() - 100)
        self.token_label.configure(wraplength=wraplength)
        self.speak_preview_label.configure(wraplength=wraplength)

    def _set_reader_controls_enabled(self, enabled: bool):
        navigation_state = "normal" if enabled else "disabled"
        speech_state = "normal" if enabled and self.engine else "disabled"
        self.previous_button.configure(state=navigation_state)
        self.reset_button.configure(state=navigation_state)
        self.speak_next_button.configure(state=speech_state)
        self.repeat_button.configure(state=speech_state)
        self.voice_menu.configure(state="normal" if self.engine else "disabled")

    def open_file(self):
        filepath = filedialog.askopenfilename(
            title="Open text file",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
        )
        if filepath:
            self.load_file(filepath)

    def load_file(self, filepath: str):
        if not os.path.exists(filepath):
            messagebox.showerror("File not found", f"File not found:\n{filepath}")
            return

        try:
            session = ReaderSession(filepath, on_changed_file=self._confirm_reset_changed_file)
        except Exception as exc:
            messagebox.showerror("Unable to open file", str(exc))
            self.status_var.set(f"Unable to open file: {exc}")
            return

        self.session = session
        self.file_var.set(filepath)
        self._set_reader_controls_enabled(True)
        engine_ready = self._ensure_engine()
        self._update_view()
        self._prime_audio_cache()
        if engine_ready:
            self.status_var.set(
                f"Loaded {len(self.session.tokens)} tokens. "
                f"Speech [{self.engine.backend.upper()}]: {self.engine.current_voice_label()}"
            )
        else:
            self.status_var.set(f"Loaded {len(self.session.tokens)} tokens, but speech is unavailable.")

    def _confirm_reset_changed_file(self) -> bool:
        return messagebox.askyesno(
            "Text changed",
            "The text file changed since the last saved progress.\n\n"
            "Reset saved progress and start from the beginning?",
            default=messagebox.YES,
        )

    def _backend_changed(self, _value):
        if self.engine:
            self._replace_engine()

    def _voice_changed(self, value):
        if self.updating_voice or not self.engine:
            return

        voice_query = self._voice_query_from_label(value)
        self._replace_engine(preferred_voice=voice_query)

    def _voice_query_from_label(self, label: str) -> str:
        prefix = label.split(" - ", 1)[0].strip()
        if prefix.startswith("en-"):
            return prefix
        return label

    def _ensure_engine(self) -> bool:
        if self.engine:
            return True
        return self._replace_engine()

    def _replace_engine(self, preferred_voice: str | None = None) -> bool:
        old_engine = self.engine
        self.status_var.set("Initializing speech engine...")

        try:
            new_engine = WindowsSpeechEngine(
                preferred_voice=preferred_voice,
                prompt_for_voice=False,
                quiet=True,
                backend_preference=self.backend_var.get(),
            )
        except Exception as exc:
            self.status_var.set(f"Speech engine error: {exc}")
            self._set_reader_controls_enabled(self.session is not None)
            return False

        self.engine = new_engine
        if old_engine:
            try:
                old_engine.shutdown()
            except Exception:
                pass

        self._refresh_voice_options()
        self._set_reader_controls_enabled(self.session is not None)
        self.status_var.set(
            f"Speech ready [{self.engine.backend.upper()}]: {self.engine.current_voice_label()}"
        )
        self._prime_audio_cache()
        return True

    def _refresh_voice_options(self):
        if not self.engine:
            return

        self.updating_voice = True
        try:
            voices = self.engine.list_voice_descriptions()
            self.voice_menu.configure(values=voices or [""])
            current_label = self.engine.current_voice_label()
            self.voice_var.set(current_label if current_label else (voices[0] if voices else ""))
        finally:
            self.updating_voice = False

    def _update_view(self, token_override=None):
        if not self.session:
            self.displayed_token = None
            self.token_label.configure(text="Open a text file")
            self.speak_preview_label.configure(text="")
            self.progress_var.set("0 / 0")
            self.progress_bar.set(0)
            return

        token = token_override or self.session.current()
        self.displayed_token = token
        self.token_label.configure(text=token.display_text)
        self.speak_preview_label.configure(text=token.speak_text if not token.is_eof else EOF_MESSAGE)
        self.progress_var.set(self.session.progress_label())

        total = len(self.session.tokens)
        self.progress_bar.set(self.session.position / total if total else 0)

    def _prime_audio_cache(self):
        if self.engine and self.session:
            self.engine.prefetch(self.session.prefetch_texts())

    def speak_next(self):
        if not self.session:
            self.status_var.set("Open a text file to begin.")
            return
        if not self._ensure_engine():
            return

        token = self.session.next()
        self._update_view(token_override=token)
        self._prime_audio_cache()
        self._speak_async(token.speak_text)

        if token.is_eof:
            self.status_var.set(EOF_MESSAGE)
        else:
            self.status_var.set(f"Speaking {token.position + 1} / {len(self.session.tokens)}")

    def repeat_token(self):
        if not self.session:
            self.status_var.set("Open a text file to begin.")
            return
        if not self._ensure_engine():
            return

        token = self.displayed_token or self.session.repeat()
        self._speak_async(token.speak_text)
        self.status_var.set("Repeating current token." if not token.is_eof else EOF_MESSAGE)

    def previous_token(self):
        if not self.session:
            self.status_var.set("Open a text file to begin.")
            return

        self.session.previous()
        self._update_view()
        self._prime_audio_cache()
        self.status_var.set("Moved back one token.")

    def reset_progress(self):
        if not self.session:
            return

        self.session.reset_progress()
        self._update_view()
        self._prime_audio_cache()
        self.status_var.set("Progress reset.")

    def _speak_async(self, text: str):
        if not self.engine or not text:
            return

        engine = self.engine
        future = self.speech_executor.submit(engine.say, text)
        future.add_done_callback(self._speech_done)

    def _speech_done(self, future):
        try:
            self.after(0, lambda: self._handle_speech_result(future))
        except RuntimeError:
            pass

    def _handle_speech_result(self, future):
        try:
            exc = future.exception()
        except concurrent.futures.CancelledError:
            return

        if exc:
            self.status_var.set(f"Speech error: {exc}")

    def close(self):
        if self.session:
            try:
                self.session.save_progress()
            except Exception:
                pass

        if self.engine:
            try:
                self.engine.shutdown()
            except Exception:
                pass

        self.speech_executor.shutdown(wait=False, cancel_futures=True)
        self.destroy()


def main():
    app = ReadWordApp()
    if len(sys.argv) > 1:
        app.after(100, lambda: app.load_file(sys.argv[1]))
    app.mainloop()


if __name__ == "__main__":
    main()
