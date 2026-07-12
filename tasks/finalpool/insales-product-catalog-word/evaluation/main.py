"""Evaluation for insales-product-catalog-word (InSales)."""
import argparse
import os
import sys
from docx import Document


# Expected catalog items and their match status
# When doing case-insensitive substring match against InSales (wc.*) products:
# NOTE: product names are KEEP-English realia (eval keys); do NOT translate.
EXPECTED_MATCHING = [
    "Canon M50 Mark II",
    "JBL Flip 4 Portable Wireless Speaker",
    "Boult Audio Powerbuds",
    "AmazonBasics Expanding File Folder",
    "CraftDev A500MB Portable Paper Trimmer",
    "AGARO Adjustable Camera Tripod Stand",
    "Ambrane Mobile Holding Tabletop Stand",
]

EXPECTED_MISSING = [
    "Belkin Ultra HD High Speed HDMI Cable",
    "Sony WH-1000XM5 Wireless Headphones",
    "Logitech MX Master 3S Mouse",
    "Samsung T7 Shield Portable SSD 1TB",
    "Anker PowerCore 26800mAh Battery Pack",
]

TOTAL_CATALOG = 12

# Heading detection: agent legitimately writes Russian headings (per task.md),
# groundtruth example uses English. Accept either language.
MATCHING_HEAD = ["matching", "совпад"]
MISSING_HEAD = ["missing", "отсутств"]
SUMMARY_HEAD = ["summary", "итог", "сводк"]


def head_matches(heading, keys):
    return any(k in heading for k in keys)


def section_text(paragraphs_by_section, keys):
    for key, paras in paragraphs_by_section.items():
        if head_matches(key, keys):
            return " ".join(paras).lower()
    return ""


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--groundtruth_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    agent_doc = os.path.join(args.agent_workspace, "Product_Comparison.docx")
    if not os.path.exists(agent_doc):
        print("FAIL: Agent output Product_Comparison.docx not found")
        sys.exit(1)

    doc = Document(agent_doc)

    # Extract all text + per-section paragraphs
    all_text = ""
    headings = []
    paragraphs_by_section = {}
    current_section = ""
    for para in doc.paragraphs:
        text = para.text.strip()
        all_text += text.lower() + " "
        if para.style.name.startswith("Heading"):
            current_section = text.lower()
            headings.append(text.lower())
            paragraphs_by_section[current_section] = []
        elif current_section and text:
            paragraphs_by_section.setdefault(current_section, []).append(text)

    matching_text = section_text(paragraphs_by_section, MATCHING_HEAD)
    missing_text = section_text(paragraphs_by_section, MISSING_HEAD)
    summary_text = section_text(paragraphs_by_section, SUMMARY_HEAD)

    # ------------------------------------------------------------------
    # CRITICAL CHECKS (semantic correctness). Any failure => sys.exit(1)
    # before the accuracy gate. These verify the CORE deliverable:
    # correct per-section classification + correct summary counts +
    # supplier/store price cross-reference. Use strict per-section
    # attribution (NOT the global all_text fallback).
    # ------------------------------------------------------------------
    critical_errors = []

    # 1) All 12 catalog product names appear, attributed to the CORRECT
    #    section (7 under Matching, 5 under Missing) -- not via all_text.
    matching_in_section = sum(
        1 for p in EXPECTED_MATCHING if p.lower() in matching_text
    )
    missing_in_section = sum(
        1 for p in EXPECTED_MISSING if p.lower() in missing_text
    )
    if matching_in_section != len(EXPECTED_MATCHING):
        critical_errors.append(
            f"CRITICAL: only {matching_in_section}/{len(EXPECTED_MATCHING)} "
            f"matching products attributed to the Matching section"
        )
    if missing_in_section != len(EXPECTED_MISSING):
        critical_errors.append(
            f"CRITICAL: only {missing_in_section}/{len(EXPECTED_MISSING)} "
            f"missing products attributed to the Missing section"
        )

    # 2) Correct classification: no EXPECTED_MISSING product leaks into the
    #    Matching section (misclassification is a hard failure).
    leaked = [p for p in EXPECTED_MISSING if p.lower() in matching_text]
    if leaked:
        critical_errors.append(
            f"CRITICAL: missing products misclassified as matching: {leaked}"
        )
    leaked2 = [p for p in EXPECTED_MATCHING if p.lower() in missing_text]
    if leaked2:
        critical_errors.append(
            f"CRITICAL: matching products misclassified as missing: {leaked2}"
        )

    # 3) Summary states the three counts correctly: total=12, matching=7,
    #    missing=5 (verify all three numbers, not just substring '12').
    s = summary_text if summary_text else all_text
    for label, num in (("total", 12), ("matching", 7), ("missing", 5)):
        if str(num) not in s:
            critical_errors.append(
                f"CRITICAL: Summary must state {label} count = {num}"
            )

    # 4) At least 5 of 7 matching products show BOTH a supplier price and a
    #    store price in their paragraph (core cross-reference deliverable).
    #    Supplier prices (USD, from PDF) and a second numeric store price
    #    must both appear in the same matching-section paragraph.
    supplier_prices = {
        "Canon M50 Mark II": "749.99",
        "JBL Flip 4 Portable Wireless Speaker": "129.99",
        "Boult Audio Powerbuds": "99.99",
        "AmazonBasics Expanding File Folder": "8.99",
        "CraftDev A500MB Portable Paper Trimmer": "9.49",
        "AGARO Adjustable Camera Tripod Stand": "29.99",
        "Ambrane Mobile Holding Tabletop Stand": "7.99",
    }
    import re
    matching_paras = []
    for key, paras in paragraphs_by_section.items():
        if head_matches(key, MATCHING_HEAD):
            matching_paras = paras
            break
    products_with_both_prices = 0
    for product, sup in supplier_prices.items():
        para = next(
            (p for p in matching_paras if product.lower() in p.lower()), None
        )
        if not para:
            continue
        if sup not in para:
            continue
        # Require a second distinct price-like number (the store price) in
        # the same paragraph beyond the supplier price token.
        nums = re.findall(r"\d[\d ., ]*\d|\d", para)
        # Count numeric tokens that look like prices (>=1 digit, may have
        # decimals); supplier price is one, store price should be another.
        distinct = set(n.replace(" ", "").replace(",", ".") for n in nums)
        # Drop the supplier price itself; need at least one other numeric.
        sup_norm = sup
        others = [n for n in distinct if sup_norm not in n and n not in sup_norm]
        if others:
            products_with_both_prices += 1
    if products_with_both_prices < 5:
        critical_errors.append(
            f"CRITICAL: only {products_with_both_prices}/7 matching products "
            f"show BOTH supplier and store prices in their paragraph (>=5 required)"
        )

    if critical_errors:
        print("=== CRITICAL CHECK FAILURE ===")
        for e in critical_errors:
            print(f"  {e}")
        sys.exit(1)

    # ------------------------------------------------------------------
    # NON-CRITICAL structural checks -> accuracy score (threshold >= 70).
    # ------------------------------------------------------------------
    all_errors = []
    total_checks = 0
    passed_checks = 0

    print("  Checking document structure...")
    has_matching = any(head_matches(h, MATCHING_HEAD) for h in headings)
    has_missing = any(head_matches(h, MISSING_HEAD) for h in headings)
    has_summary = any(head_matches(h, SUMMARY_HEAD) for h in headings)
    for ok, name in (
        (has_matching, "Matching heading"),
        (has_missing, "Missing heading"),
        (has_summary, "Summary heading"),
    ):
        total_checks += 1
        if ok:
            passed_checks += 1
        else:
            all_errors.append(f"Missing {name}")

    print("  Checking matching products...")
    found_matching = 0
    for product in EXPECTED_MATCHING:
        if product.lower() in matching_text or product.lower() in all_text:
            found_matching += 1
        else:
            all_errors.append(f"Matching product not found: {product}")
    total_checks += 1
    if found_matching >= 5:
        passed_checks += 1
        print(f"    {found_matching}/7 matching products found")
    else:
        all_errors.append(f"Only {found_matching}/7 matching products found")

    print("  Checking missing products...")
    found_missing = 0
    for product in EXPECTED_MISSING:
        if product.lower() in missing_text or product.lower() in all_text:
            found_missing += 1
        else:
            all_errors.append(f"Missing product not listed: {product}")
    total_checks += 1
    if found_missing >= 3:
        passed_checks += 1
        print(f"    {found_missing}/5 missing products found")
    else:
        all_errors.append(f"Only {found_missing}/5 missing products found")

    print("  Checking summary...")
    total_checks += 1
    if "12" in (summary_text if summary_text else all_text):
        passed_checks += 1
    else:
        all_errors.append("Summary should mention 12 total catalog products")

    accuracy = (passed_checks / total_checks * 100) if total_checks else 0.0
    print(f"\nAccuracy: {accuracy:.1f}% ({passed_checks}/{total_checks})")
    if all_errors:
        for e in all_errors[:15]:
            print(f"  {e}")

    if accuracy >= 70:
        print("\n=== RESULT: PASS ===")
        sys.exit(0)
    else:
        print("\n=== RESULT: FAIL ===")
        sys.exit(1)


if __name__ == "__main__":
    main()
