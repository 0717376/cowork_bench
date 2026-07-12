"""
Evaluation script for academic-rag-lit-review-ru-arxiv-word-memory task.
Checks that Literature_Review.docx exists and contains the expected content.

Usage:
  python -m evaluation.main --agent_workspace <path> --groundtruth_workspace <path> --launch_time <time>
"""
import argparse
import os
import re
import sys


PASS_COUNT = 0
FAIL_COUNT = 0
FAILED_NAMES = []


CRITICAL_CHECKS = {
    "All 5 paper titles present in document",
    "All 5 first authors present in document",
    "Document has both introduction and conclusion sections",
}


def check(name: str, condition: bool, detail: str = ""):
    global PASS_COUNT, FAIL_COUNT
    if condition:
        PASS_COUNT += 1
        print(f"  [PASS] {name}")
    else:
        FAIL_COUNT += 1
        FAILED_NAMES.append(name)
        detail_truncated = (detail[:200] + "...") if len(detail) > 200 else detail
        print(f"  [FAIL] {name}: {detail_truncated}")


def normalize(text: str) -> str:
    """Normalize text for comparison: lowercase, collapse whitespace."""
    return re.sub(r'\s+', ' ', text.lower().strip())


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", type=str, required=True)
    parser.add_argument("--groundtruth_workspace", type=str, required=True)
    parser.add_argument("--launch_time", type=str, required=False)
    parser.add_argument("--res_log_file", type=str, required=False)
    args = parser.parse_args()

    docx_path = os.path.join(args.agent_workspace, "Literature_Review.docx")

    # Check 1: File exists
    check("Literature_Review.docx exists", os.path.exists(docx_path),
          f"File not found at {docx_path}")

    if not os.path.exists(docx_path):
        print(f"\nResults: {PASS_COUNT}/{PASS_COUNT + FAIL_COUNT} passed, {FAIL_COUNT} failed")
        sys.exit(1)

    # Read the Word document
    try:
        from docx import Document
        doc = Document(docx_path)
        full_text = "\n".join([para.text for para in doc.paragraphs])
    except Exception as e:
        check("Word document readable", False, str(e))
        print(f"\nResults: {PASS_COUNT}/{PASS_COUNT + FAIL_COUNT} passed, {FAIL_COUNT} failed")
        sys.exit(1)

    normalized = normalize(full_text)

    # Check 2: Minimum content length
    check("Document has at least 500 characters",
          len(full_text.strip()) >= 500,
          f"Document has {len(full_text.strip())} characters")

    # Check 3: All 5 paper titles appear (case-insensitive partial match).
    # Titles must stay in original English so they match what preprocess
    # writes into arxiv/scholarly schemas (and what the agent reads from there).
    paper_titles = [
        "Retrieval-Augmented Generation for Knowledge-Intensive NLP Tasks",
        "Retrieval-Augmented Generation for Large Language Models: A Survey",
        "Self-RAG: Learning to Retrieve, Generate, and Critique through Self-Reflection",
        "RAPTOR: Recursive Abstractive Processing for Tree-Organized Retrieval",
        "From RAG to Rich",
    ]
    titles_found = 0
    for title in paper_titles:
        title_lower = title.lower()
        present = title_lower in normalized
        if present:
            titles_found += 1
        check(f"Paper title present: {title[:60]}...",
              present,
              "Title not found in document text")

    check("All 5 paper titles present in document",
          titles_found == 5,
          f"Only {titles_found}/5 paper titles found")

    # Check 4: All 5 first authors appear
    first_authors = ["Lewis", "Gao", "Asai", "Sarthi", "Chen"]
    authors_found = 0
    for author in first_authors:
        present = author.lower() in normalized
        if present:
            authors_found += 1
        check(f"Author present: {author}",
              present,
              f"Author '{author}' not found in document text")

    check("All 5 first authors present in document",
          authors_found == 5,
          f"Only {authors_found}/5 first authors found")

    # Check 5: Key domain terms present
    key_terms = ["retrieval", "generation", "augmented"]
    for term in key_terms:
        check(f"Key term present: {term}",
              term.lower() in normalized,
              f"Term '{term}' not found")

    # Check 6: Document has structure (introduction/conclusion).
    # Accept both English markers and Russian equivalents because the body
    # text is expected to be in Russian.
    intro_markers = ["introduction", "введение"]
    conclusion_markers = ["conclusion", "summary", "synthesis",
                          "заключение", "выводы", "итог"]
    has_intro = any(m in normalized for m in intro_markers)
    has_conclusion = any(m in normalized for m in conclusion_markers)
    check("Document has introduction section", has_intro,
          "No 'introduction' / 'введение' found in text")
    check("Document has conclusion/summary section", has_conclusion,
          "No 'conclusion' / 'заключение' / 'выводы' found in text")

    check("Document has both introduction and conclusion sections",
          has_intro and has_conclusion,
          f"intro={has_intro}, conclusion={has_conclusion}")

    # Summary
    total = PASS_COUNT + FAIL_COUNT
    critical_failed = [n for n in FAILED_NAMES if n in CRITICAL_CHECKS]
    print(f"\nResults: {PASS_COUNT}/{total} passed, {FAIL_COUNT} failed")
    if critical_failed:
        print(f"CRITICAL FAILURES: {len(critical_failed)}")
        for n in critical_failed:
            print(f"  - {n}")
        sys.exit(1)

    sys.exit(0 if FAIL_COUNT == 0 else 1)


if __name__ == "__main__":
    main()
