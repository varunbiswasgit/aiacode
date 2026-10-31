#!/usr/bin/env python3
"""Generate the RAG prompt templates (Q&A, gap analysis, suggestion/redline).

Tracker task: P5-07 (prompt templates)
Prompts: organisation_name, answer_tone, abstain_message, prompts_directory
Prerequisites: none. Review the generated templates with QA before use;
  every template forces citations and forbids inventing content.
Output: qa.txt, gap.txt, suggest.txt in the prompts directory.
"""
import os


def prompt(label, default=""):
    while True:
        value = input(f"{label} [{default}]: ").strip() or default
        if value:
            return value
        print("A value is required.")


def main():
    print("=== Step 1: Template settings ===")
    organisation_name = prompt("Organisation name (example — replace with your value)", "Internal")
    answer_tone = prompt("Answer tone (concise/detailed) (example — replace with your value)", "concise")
    abstain_message = prompt("Message when the answer is not in the documents (example — replace with your value)",
                             "Not found in internal documents.")
    prompts_directory = prompt("Prompts directory (example — replace with your value)", "/opt/airag/prompts")
    os.makedirs(prompts_directory, exist_ok=True)

    qa_template = f"""You are the {organisation_name} internal document assistant.
Answer ONLY from the context chunks below. Be {answer_tone}.
Every factual claim must cite its source as [doc_id#chunk_index].
If the context does not support an answer, reply exactly: {abstain_message}
Never invent clauses, numbers or document content.

Context:
{{context}}

Question: {{question}}

Answer (with citations):
"""

    gap_template = f"""You are a {organisation_name} quality reviewer.
Compare the NEW standard clause below against the internal document section.
Report: (1) impact rating High/Medium/Low, (2) exactly what is now out of
alignment, quoting both texts, (3) a suggested revised wording for the
internal section. Do NOT edit any document yourself - you only suggest.
Cite the clause (standard, clause_no, version) for every finding.

New clause:
{{new_clause}}

Internal document section:
{{doc_section}}

Gap analysis:
"""

    suggest_template = f"""You are a {organisation_name} document improvement assistant.
Given the gap analysis below, write a redline-style suggestion: the current
wording, the proposed wording, and the reason tied to the clause.
The suggestion goes to a human reviewer for approval and is NEVER applied
automatically to a controlled document.

Gap analysis:
{{gap_analysis}}

Redline suggestion:
"""

    for filename, content in (("qa.txt", qa_template), ("gap.txt", gap_template),
                              ("suggest.txt", suggest_template)):
        path = os.path.join(prompts_directory, filename)
        with open(path, "w", encoding="utf-8") as handle:
            handle.write(content)
        print(f"  Wrote {path}")
    print("Done. Review these templates with QA, then commit them to the repo.")


if __name__ == "__main__":
    main()
