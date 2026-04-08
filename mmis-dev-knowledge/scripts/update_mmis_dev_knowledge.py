from __future__ import annotations

import argparse
import json
import sys
from datetime import date
from pathlib import Path
from typing import Any


BASE_DIR = Path(__file__).resolve().parent.parent
KNOWLEDGE_FILE = BASE_DIR / "references" / "knowledge-base.json"
VALID_SECTIONS = {
    "loginStrategies",
    "apiPatterns",
    "commonErrors",
    "performanceTips",
    "stableWorkflows",
    "toolingNotes",
}
VALID_STATUS = {"confirmed", "experimental", "deprecated"}


def load_knowledge() -> dict[str, Any]:
    return json.loads(KNOWLEDGE_FILE.read_text(encoding="utf-8"))


def save_knowledge(data: dict[str, Any]) -> None:
    data.setdefault("meta", {})
    data["meta"]["last_updated"] = str(date.today())
    KNOWLEDGE_FILE.write_text(
        json.dumps(data, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )


def normalize_entry(entry_id: str, payload: dict[str, Any]) -> dict[str, Any]:
    required = {
        "title",
        "description",
        "applicable_when",
        "advantages",
        "disadvantages",
        "recommended_when",
        "status",
        "reusable",
        "notes"
    }
    missing = sorted(required - payload.keys())
    if missing:
        raise ValueError(f"entry 缺少欄位: {', '.join(missing)}")
    if payload["status"] not in VALID_STATUS:
        raise ValueError(f"status 必須是 {sorted(VALID_STATUS)} 之一")

    normalized = {
        "id": entry_id,
        "title": str(payload["title"]).strip(),
        "description": str(payload["description"]).strip(),
        "applicable_when": sorted(set(map(str, payload["applicable_when"]))),
        "advantages": sorted(set(map(str, payload["advantages"]))),
        "disadvantages": sorted(set(map(str, payload["disadvantages"]))),
        "recommended_when": sorted(set(map(str, payload["recommended_when"]))),
        "status": payload["status"],
        "reusable": bool(payload["reusable"]),
        "notes": sorted(set(map(str, payload.get("notes", []))))
    }
    replaces = payload.get("replaces", [])
    if replaces:
        normalized["replaces"] = sorted(set(map(str, replaces)))
    return normalized


def merge_entries(existing: dict[str, Any], incoming: dict[str, Any]) -> dict[str, Any]:
    merged = dict(existing)
    for key in ["title", "description", "status", "reusable"]:
        merged[key] = incoming[key]
    for key in ["applicable_when", "advantages", "disadvantages", "recommended_when", "notes"]:
        merged[key] = sorted(set(existing.get(key, []) + incoming.get(key, [])))
    if incoming.get("replaces") or existing.get("replaces"):
        merged["replaces"] = sorted(set(existing.get("replaces", []) + incoming.get("replaces", [])))
    return merged


def deprecate_replaced_entries(section_items: list[dict[str, Any]], entry: dict[str, Any]) -> list[str]:
    deprecated: list[str] = []
    for replaced_id in entry.get("replaces", []):
        for item in section_items:
            if item.get("id") == replaced_id:
                item["status"] = "deprecated"
                notes = set(item.get("notes", []))
                notes.add(f"Superseded by {entry['id']}.")
                item["notes"] = sorted(notes)
                deprecated.append(replaced_id)
    return deprecated


def update_knowledge(section: str, entry_id: str, payload: dict[str, Any]) -> dict[str, Any]:
    if section not in VALID_SECTIONS:
        raise ValueError(f"section 必須是 {sorted(VALID_SECTIONS)} 之一")

    knowledge = load_knowledge()
    section_items = knowledge.setdefault(section, [])
    normalized = normalize_entry(entry_id, payload)

    action = "added"
    for index, item in enumerate(section_items):
        if item.get("id") == entry_id:
            section_items[index] = merge_entries(item, normalized)
            normalized = section_items[index]
            action = "merged"
            break
    else:
        section_items.append(normalized)

    deprecated = deprecate_replaced_entries(section_items, normalized)
    section_items.sort(key=lambda item: item["id"])
    save_knowledge(knowledge)
    return {
        "section": section,
        "id": entry_id,
        "action": action,
        "deprecated": deprecated,
        "path": str(KNOWLEDGE_FILE)
    }


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Update structured MMIS development knowledge")
    parser.add_argument("--input", required=True, help="Path to a JSON payload file")
    return parser


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")

    args = build_parser().parse_args()
    payload_path = Path(args.input)
    payload = json.loads(payload_path.read_text(encoding="utf-8"))
    result = update_knowledge(
        section=payload["section"],
        entry_id=payload["id"],
        payload=payload["entry"]
    )
    print(json.dumps(result, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
