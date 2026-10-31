#!/usr/bin/env python3
"""Filter retrieval candidates by the user's AD groups BEFORE generation.

Tracker task: P5-10 (permission filter)
Prompts: username, user_ad_groups (comma-separated), access_map_path,
         candidates_json_path, output_json_path
Prerequisites: candidates.json from 05-hybrid-retrieve.py; an access map
  JSON of the form {"AD-GROUP": ["doc-prefix-1", "doc-prefix-2"], ...}.
  A chunk is kept only if its doc_id starts with one of the prefixes of
  at least one of the user's groups. If a group is missing from the map,
  it grants NOTHING (fail closed).
Output: filtered candidates JSON. Apply this BEFORE reranking/generation.
"""
import json


def prompt(label, default=""):
    while True:
        value = input(f"{label} [{default}]: ").strip() or default
        if value:
            return value
        print("A value is required.")


def main():
    print("=== Step 1: Filter settings ===")
    username = prompt("Username")
    user_ad_groups = prompt("User's AD groups, comma-separated", "AIRAG-Users")
    access_map_path = prompt("Access map JSON path", "/opt/airag/access-map.json")
    candidates_json_path = prompt("Candidates JSON", "candidates.json")
    output_json_path = prompt("Filtered output JSON", "candidates-filtered.json")

    with open(access_map_path, encoding="utf-8") as handle:
        access_map = json.load(handle)
    allowed_prefixes = []
    for group in [entry.strip() for entry in user_ad_groups.split(",") if entry.strip()]:
        allowed_prefixes.extend(access_map.get(group, []))
    if not allowed_prefixes:
        print(f"WARNING: user {username} has no mapped groups - everything is denied (fail closed).")

    with open(candidates_json_path, encoding="utf-8") as handle:
        candidates = json.load(handle)
    kept = [candidate for candidate in candidates
            if any(candidate.get("doc_id", "").startswith(prefix)
                   for prefix in allowed_prefixes)]
    with open(output_json_path, "w", encoding="utf-8") as handle:
        json.dump(kept, handle, indent=2)
    print(f"Done. {len(kept)}/{len(candidates)} candidates kept for {username}.")
    print("Feed the filtered file (not the original) into 05-rerank-results.py.")


if __name__ == "__main__":
    main()
