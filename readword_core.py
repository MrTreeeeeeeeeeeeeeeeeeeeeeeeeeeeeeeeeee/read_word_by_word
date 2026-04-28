import hashlib
import json
import os
import re
import unicodedata
from dataclasses import dataclass
from datetime import datetime
from typing import Callable, List, Optional, Tuple


TOKEN_MAP = {
    "\n": ("<NEW LINE>", "new line"),
    "...": ("<ELLIPSIS>", "ellipsis"),
    "…": ("<ELLIPSIS>", "ellipsis"),
    ".": ("<PERIOD>", "period"),
    ",": ("<COMMA>", "comma"),
    ":": ("<COLON>", "colon"),
    ";": ("<SEMICOLON>", "semicolon"),
    "!": ("<EXCLAMATION>", "exclamation mark"),
    "?": ("<QUESTION>", "question mark"),
    '"': ("<QUOTE>", "quote"),
    "“": ("<QUOTE>", "quote"),
    "”": ("<QUOTE>", "quote"),
    "-": ("<DASH>", "dash"),
    "—": ("<EM DASH>", "dash"),
}
SYMBOL_NAME_MAP = {
    "(": ("<LEFT PARENTHESIS>", "left parenthesis"),
    ")": ("<RIGHT PARENTHESIS>", "right parenthesis"),
    "[": ("<LEFT BRACKET>", "left bracket"),
    "]": ("<RIGHT BRACKET>", "right bracket"),
    "{": ("<LEFT BRACE>", "left brace"),
    "}": ("<RIGHT BRACE>", "right brace"),
    "<": ("<LESS THAN>", "less than"),
    ">": ("<GREATER THAN>", "greater than"),
    "/": ("<SLASH>", "slash"),
    "\\": ("<BACKSLASH>", "backslash"),
    "|": ("<VERTICAL BAR>", "vertical bar"),
    "_": ("<UNDERSCORE>", "underscore"),
    "=": ("<EQUALS>", "equals sign"),
    "+": ("<PLUS>", "plus sign"),
    "*": ("<ASTERISK>", "asterisk"),
    "&": ("<AMPERSAND>", "ampersand"),
    "@": ("<AT>", "at sign"),
    "#": ("<HASH>", "hash sign"),
    "$": ("<DOLLAR>", "dollar sign"),
    "%": ("<PERCENT>", "percent sign"),
    "^": ("<CARET>", "caret"),
    "~": ("<TILDE>", "tilde"),
    "`": ("<BACKTICK>", "backtick"),
}
CARDINAL_NUMBER_WORDS = {
    0: "zero",
    1: "one",
    2: "two",
    3: "three",
    4: "four",
    5: "five",
    6: "six",
    7: "seven",
    8: "eight",
    9: "nine",
    10: "ten",
    11: "eleven",
    12: "twelve",
}
ORDINAL_NUMBER_WORDS = {
    1: "first",
    2: "second",
    3: "third",
    4: "fourth",
    5: "fifth",
    6: "sixth",
    7: "seventh",
    8: "eighth",
    9: "ninth",
    10: "tenth",
    11: "eleventh",
    12: "twelfth",
    13: "thirteenth",
    14: "fourteenth",
    15: "fifteenth",
    16: "sixteenth",
    17: "seventeenth",
    18: "eighteenth",
    19: "nineteenth",
    20: "twentieth",
    21: "twenty first",
    22: "twenty second",
    23: "twenty third",
    24: "twenty fourth",
    25: "twenty fifth",
    26: "twenty sixth",
    27: "twenty seventh",
    28: "twenty eighth",
    29: "twenty ninth",
    30: "thirtieth",
    31: "thirty first",
}
SINGULAR_ORDINAL_DENOMINATORS = {
    2: "half",
    3: "third",
    4: "fourth",
    5: "fifth",
    6: "sixth",
    7: "seventh",
    8: "eighth",
    9: "ninth",
    10: "tenth",
    11: "eleventh",
    12: "twelfth",
}

EDGE_PREFETCH_TOKENS = 24
EOF_MESSAGE = "End of file reached."
TOKEN_PATTERN = (
    r"(?:[A-Za-z]\.){2,}[A-Za-z]?\.?"
    r"|\d{4}-\d{4}s?"
    r"|\d+\s*(?:st|nd|rd|th)"
    r"|\d+/\d+(?:\s*th)?"
    r"|[Vv]s\."
    r"|[Vv]\."
    r"|\.{3}"
    r"|[\w'’]+"
    r"|[^\w\s]"
    r"|\n"
)


@dataclass(frozen=True)
class ReaderToken:
    position: int
    token: str
    display_text: str
    speak_text: str
    is_eof: bool = False


def tokenize_text(text: str) -> List[str]:
    return [token for token in re.findall(TOKEN_PATTERN, text) if token.strip() or token == "\n"]


def hash_text(text: str) -> str:
    return hashlib.sha1(text.encode("utf-8")).hexdigest()


def progress_path_for(filepath: str) -> str:
    return os.path.splitext(filepath)[0] + "_progress.json"


def speak_four_digit_number(token: str) -> str:
    first_half = token[:2]
    second_half = token[2:]
    if second_half.startswith("0"):
        return f"{first_half} oh {int(second_half)}"
    return f"{first_half} {second_half}"


def get_display_info(token: str) -> Tuple[str, str]:
    if token in TOKEN_MAP:
        return TOKEN_MAP[token]

    if token in SYMBOL_NAME_MAP:
        return SYMBOL_NAME_MAP[token]

    dotted_acronym_match = re.fullmatch(r"(?:[A-Za-z]\.){2,}[A-Za-z]?\.?", token)
    if dotted_acronym_match:
        letters = [char for char in token if char.isalpha()]
        return (token, " ".join(letter.upper() for letter in letters))

    year_range_match = re.fullmatch(r"(\d{4})-(\d{4}s?)", token)
    if year_range_match:
        start_year = year_range_match.group(1)
        end_year = year_range_match.group(2)
        return (token, f"{speak_four_digit_number(start_year)} to {end_year}")

    ordinal_match = re.fullmatch(r"(\d+)\s*(st|nd|rd|th)", token, flags=re.IGNORECASE)
    if ordinal_match:
        number = int(ordinal_match.group(1))
        ordinal_word = ORDINAL_NUMBER_WORDS.get(number)
        if ordinal_word:
            return (token, ordinal_word)
        return (token, token.replace(" ", ""))

    if re.fullmatch(r"\d+/\d+(?:\s*th)?", token):
        numerator_text, denominator_text = token.split("/", 1)
        denominator_text = denominator_text.replace(" ", "")
        denominator_text = re.sub(r"th$", "", denominator_text)
        numerator = int(numerator_text)
        denominator = int(denominator_text)

        numerator_word = CARDINAL_NUMBER_WORDS.get(numerator, str(numerator))
        denominator_word = SINGULAR_ORDINAL_DENOMINATORS.get(denominator)

        if denominator_word:
            if numerator != 1:
                if denominator_word.endswith("f"):
                    denominator_word = denominator_word[:-1] + "ves"
                else:
                    denominator_word += "s"
            return (token, f"{numerator_word} {denominator_word}")

        return (token, f"{numerator} over {denominator}")

    if token.casefold() in {"v.", "vs", "vs."}:
        return (token, "versus")

    if token.isdigit() and len(token) == 4:
        return (token, speak_four_digit_number(token))

    if token.isalpha() and token.isupper() and len(token) > 1:
        return (token, " ".join(token))

    if len(token) == 1:
        category = unicodedata.category(token)
        if category.startswith(("P", "S")):
            symbol_name = unicodedata.name(token, "")
            if symbol_name:
                friendly_name = symbol_name.lower().replace("-", " ")
                return (f"<{symbol_name}>", friendly_name)

    return (token, token)


class ReaderSession:
    def __init__(
        self,
        filepath: str,
        on_changed_file: Optional[Callable[[], bool]] = None,
    ):
        self.filepath = filepath
        self.save_file = progress_path_for(filepath)
        self.on_changed_file = on_changed_file
        self.source_text = self._read_file_text()
        self.source_hash = hash_text(self.source_text)
        self.tokens = tokenize_text(self.source_text)
        self.position = self.load_progress()

    def _read_file_text(self) -> str:
        with open(self.filepath, "r", encoding="utf-8") as file:
            return file.read()

    def _clamp_position(self, position: int) -> int:
        return max(0, min(position, len(self.tokens)))

    def _write_progress_file(self, position: int):
        with open(self.save_file, "w", encoding="utf-8") as file:
            json.dump(
                {
                    "position": self._clamp_position(position),
                    "timestamp": str(datetime.now()),
                    "source_hash": self.source_hash,
                },
                file,
            )

    def load_progress(self) -> int:
        if not os.path.exists(self.save_file):
            return 0

        try:
            with open(self.save_file, "r", encoding="utf-8") as file:
                data = json.load(file)
        except Exception:
            return 0

        try:
            position = self._clamp_position(int(data.get("position", 0)))
        except (TypeError, ValueError):
            position = 0
        saved_hash = data.get("source_hash")

        if saved_hash and saved_hash != self.source_hash:
            should_reset = True if self.on_changed_file is None else self.on_changed_file()
            if should_reset:
                self._write_progress_file(0)
                return 0

            self._write_progress_file(position)

        return position

    def save_progress(self):
        self._write_progress_file(self.position)

    def reset_progress(self):
        self.position = 0
        self.save_progress()
        return self.current()

    def current(self) -> ReaderToken:
        if not self.tokens or self.position >= len(self.tokens):
            return ReaderToken(
                position=self.position,
                token="",
                display_text=EOF_MESSAGE,
                speak_text=EOF_MESSAGE,
                is_eof=True,
            )

        token = self.tokens[self.position]
        display_text, speak_text = get_display_info(token)
        return ReaderToken(
            position=self.position,
            token=token,
            display_text=display_text,
            speak_text=speak_text,
        )

    def next(self) -> ReaderToken:
        token = self.current()
        if not token.is_eof:
            self.position = self._clamp_position(self.position + 1)
            self.save_progress()
        return token

    def previous(self) -> ReaderToken:
        if self.position > 0:
            self.position = self._clamp_position(self.position - 1)
            self.save_progress()
        return self.current()

    def repeat(self) -> ReaderToken:
        return self.current()

    def jump_to(self, position: int) -> ReaderToken:
        self.position = self._clamp_position(position)
        self.save_progress()
        return self.current()

    def progress_label(self) -> str:
        return f"{self.position} / {len(self.tokens)}"

    def prefetch_texts(self, lookahead: int = EDGE_PREFETCH_TOKENS) -> List[str]:
        texts = []
        end = min(len(self.tokens), self.position + lookahead)
        for index in range(self.position, end):
            _, speak_text = get_display_info(self.tokens[index])
            texts.append(speak_text)
        return texts
