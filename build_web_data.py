"""Create the compact static-data payload consumed by the mobile PWA."""

from __future__ import annotations

import csv
import json
import re
from pathlib import Path


ROOT = Path(__file__).resolve().parent
SOURCE = ROOT / "build" / "oxford_5000_entries.csv"
DESTINATION = ROOT / "web_app" / "data" / "words.json"

POS_LABELS = {
    "n.": "명사",
    "v.": "동사",
    "adj.": "형용사",
    "adv.": "부사",
    "prep.": "전치사",
    "conj.": "접속사",
    "pron.": "대명사",
    "det.": "한정사",
    "exclam.": "감탄사",
    "number": "수사",
}


def korean_pos(source_detail: str) -> str:
    cleaned = re.sub(r"\b(?:B1|B2|C1)\b", "", source_detail)
    tokens = re.findall(
        r"n\.|v\.|adj\.|adv\.|prep\.|conj\.|pron\.|det\.|exclam\.|number",
        cleaned,
    )
    if not tokens:
        raise ValueError(f"Unknown part-of-speech detail: {source_detail}")
    return " / ".join(POS_LABELS[token] for token in tokens)


def main() -> None:
    with SOURCE.open(encoding="utf-8-sig", newline="") as file:
        source_rows = list(csv.DictReader(file))

    words = [
        {
            "id": f"{index:04d}",
            "word": row["word"].strip(),
            "partOfSpeech": korean_pos(row["source_detail"]),
            "meaning": row["meaning"].strip(),
            "letter": row["word"].strip()[0].upper(),
        }
        for index, row in enumerate(source_rows, start=1)
    ]

    if len(words) != 2000:
        raise ValueError(f"Expected 2,000 entries, got {len(words)}.")
    if len({word["id"] for word in words}) != len(words):
        raise ValueError("Word IDs must be unique.")
    if len({(word["word"], word["partOfSpeech"], word["meaning"]) for word in words}) != len(words):
        raise ValueError("Duplicate word entries were found.")
    if any(not word["word"] or not word["partOfSpeech"] or not word["meaning"] for word in words):
        raise ValueError("Every entry must have a word, part of speech, and meaning.")

    DESTINATION.parent.mkdir(parents=True, exist_ok=True)
    DESTINATION.write_text(
        json.dumps(words, ensure_ascii=False, separators=(",", ":")),
        encoding="utf-8",
    )
    print(f"Wrote {len(words)} entries to {DESTINATION}")


if __name__ == "__main__":
    main()
