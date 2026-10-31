#!/usr/bin/env python3
"""Evaluate RAG answers against the golden set: retrieval hit rate,
citation accuracy, and abstain behaviour.

Tracker task: P9-02 (run evaluation)
Prompts: golden_json_path, answers_jsonl_path, output_report_path
Prerequisites: golden.json from 09-create-golden-set.py and an answers
  JSONL produced by running the app (or 05-generate-answer.py) over every
  golden question. Each answers line: {question, answer_text,
  cited_doc_ids, abstained}.
Metrics:
  - retrieval hit rate: % of questions where at least one cited doc is an
    expected doc.
  - citation accuracy: % of non-abstained answers that cite at least one doc.
  - abstain rate: % of questions the system refused to answer.
  - keyword overlap: % of expected-answer words present in the answer
    (a rough faithfulness proxy; read failures manually, do not trust it blindly).
Output: a Markdown report. Fix the top failure modes, then re-run (P9-03).
"""
import json
import re
import string


def prompt(label, default=""):
    while True:
        value = input(f"{label} [{default}]: ").strip() or default
        if value:
            return value
        print("A value is required.")


def tokenize(text):
    translator = str.maketrans("", "", string.punctuation)
    return set(text.lower().translate(translator).split())


def main():
    print("=== Step 1: Evaluation inputs ===")
    golden_json_path = prompt("Golden set JSON", "/opt/airag/eval/golden.json")
    answers_jsonl_path = prompt("Answers JSONL", "/tmp/airag-answers.jsonl")
    output_report_path = prompt("Report output path", "/opt/airag/eval/eval-report.md")

    with open(golden_json_path, encoding="utf-8") as handle:
        golden = json.load(handle)
    answers = {}
    with open(answers_jsonl_path, encoding="utf-8") as handle:
        for line in handle:
            record = json.loads(line)
            answers[record["question"]] = record

    print("=== Step 2: Scoring ===")
    total = len(golden)
    hits = 0
    cited = 0
    abstains = 0
    overlaps = []
    failures = []
    for item in golden:
        answer = answers.get(item["question"])
        if answer is None:
            failures.append((item["question"], "no answer recorded"))
            continue
        if answer.get("abstained"):
            abstains += 1
            continue
        cited_docs = set(answer.get("cited_doc_ids", []))
        if cited_docs:
            cited += 1
        if cited_docs & set(item.get("expected_doc_ids", [])):
            hits += 1
        else:
            failures.append((item["question"], "no expected doc cited"))
        expected_words = tokenize(item.get("expected_answer", ""))
        answer_words = tokenize(answer.get("answer_text", ""))
        if expected_words:
            overlaps.append(len(expected_words & answer_words) / len(expected_words))

    answered = total - abstains
    report = (
        "# AI-RAG evaluation report\n\n"
        f"- Golden questions: {total}\n"
        f"- Retrieval hit rate: {hits}/{total} ({100 * hits // total if total else 0}%)\n"
        f"- Citation accuracy (of answered): {cited}/{answered} "
        f"({100 * cited // answered if answered else 0}%)\n"
        f"- Abstain rate: {abstains}/{total} ({100 * abstains // total if total else 0}%)\n"
        f"- Mean expected-answer keyword overlap: "
        f"{sum(overlaps) / len(overlaps):.2f}\n\n"
        "## Failures to fix\n\n"
        + "".join(f"- {question} ({reason})\n" for question, reason in failures)
    )
    with open(output_report_path, "w", encoding="utf-8") as handle:
        handle.write(report)
    print(report)
    print(f"Report saved to {output_report_path}.")
    print("Next: fix the top failure causes (chunking, rerank, prompts) and re-run.")


if __name__ == "__main__":
    main()
