import json

from readword_core import (
    ReaderSession,
    get_display_info,
    hash_text,
    tokenize_text,
)


def write_text(path, text):
    path.write_text(text, encoding="utf-8")


def write_progress(path, position, source_hash):
    path.write_text(
        json.dumps(
            {
                "position": position,
                "timestamp": "2026-04-27 12:00:00",
                "source_hash": source_hash,
            }
        ),
        encoding="utf-8",
    )


def test_tokenize_text_keeps_special_word_tokens_and_punctuation():
    text = "Revolution vs. U.S.A. v. 21st 1/2 1776-1780s!\nNext"

    assert tokenize_text(text) == [
        "Revolution",
        "vs.",
        "U.S.A.",
        "v.",
        "21st",
        "1/2",
        "1776-1780s",
        "!",
        "\n",
        "Next",
    ]


def test_get_display_info_maps_punctuation_numbers_acronyms_and_fractions():
    assert get_display_info(".") == ("<PERIOD>", "period")
    assert get_display_info("1776") == ("1776", "17 76")
    assert get_display_info("USA") == ("USA", "U S A")
    assert get_display_info("1/2") == ("1/2", "one half")
    assert get_display_info("3/4") == ("3/4", "three fourths")


def test_reader_session_saves_and_loads_progress(tmp_path):
    text_file = tmp_path / "lesson.txt"
    write_text(text_file, "one two three")

    session = ReaderSession(str(text_file))
    session.jump_to(2)

    saved = json.loads((tmp_path / "lesson_progress.json").read_text(encoding="utf-8"))
    assert saved["position"] == 2
    assert saved["source_hash"] == hash_text("one two three")
    assert saved["timestamp"]

    loaded = ReaderSession(str(text_file))
    assert loaded.position == 2
    assert loaded.current().token == "three"


def test_reader_session_clamps_out_of_range_progress(tmp_path):
    text_file = tmp_path / "lesson.txt"
    write_text(text_file, "one two three")
    write_progress(tmp_path / "lesson_progress.json", 99, hash_text("one two three"))

    session = ReaderSession(str(text_file))

    assert session.position == 3
    assert session.current().is_eof


def test_reader_session_resets_changed_text_progress(tmp_path):
    text_file = tmp_path / "lesson.txt"
    write_text(text_file, "new text")
    write_progress(tmp_path / "lesson_progress.json", 1, "old-hash")

    calls = []
    session = ReaderSession(str(text_file), on_changed_file=lambda: calls.append(True) or True)

    saved = json.loads((tmp_path / "lesson_progress.json").read_text(encoding="utf-8"))
    assert calls == [True]
    assert session.position == 0
    assert saved["position"] == 0
    assert saved["source_hash"] == hash_text("new text")


def test_reader_session_keeps_changed_text_progress_when_requested(tmp_path):
    text_file = tmp_path / "lesson.txt"
    write_text(text_file, "new text")
    write_progress(tmp_path / "lesson_progress.json", 99, "old-hash")

    calls = []
    session = ReaderSession(str(text_file), on_changed_file=lambda: calls.append(False) or False)

    saved = json.loads((tmp_path / "lesson_progress.json").read_text(encoding="utf-8"))
    assert calls == [False]
    assert session.position == 2
    assert saved["position"] == 2
    assert saved["source_hash"] == hash_text("new text")
