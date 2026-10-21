"""Bounded, parameterized compiler for the query syntax advertised by this mock.

OR, AND/implicit AND, ANDNOT, parentheses, and ti:/abs:/au: fields are
interpreted before PostgreSQL FTS. Quoted phrases are literal contiguous text,
not an AND bag of words. Invalid syntax is rejected instead of broadening it.
"""

from dataclasses import dataclass
import json
import re

_FIELDS = {
    "ti": "coalesce(title, '')",
    "abs": "coalesce(summary, '')",
    "au": "coalesce(authors::text, '')",
    "all": "coalesce(title, '') || ' ' || coalesce(summary, '') || ' ' || coalesce(authors::text, '')",
}


@dataclass(frozen=True)
class SearchQuery:
    predicate: str
    params: tuple
    rank: str
    rank_params: tuple
    plain_text: str | None


def _tokens(query):
    if len(query) > 4096:
        raise ValueError("Search query exceeds 4096 characters")
    result, pos = [], 0
    while pos < len(query):
        if query[pos].isspace():
            pos += 1
            continue
        char = query[pos]
        if char in "()":
            result.append((char, char))
            pos += 1
        elif char == '"':
            match = re.match(r'"(?:[^"\\]|\\.)*"', query[pos:])
            if not match:
                raise ValueError("Unclosed quoted phrase")
            text = json.loads(match.group())
            if not text.strip():
                raise ValueError("Empty quoted phrase")
            result.append(("PHRASE", text))
            pos += len(match.group())
        else:
            field = re.match(r"([A-Za-z]+):", query[pos:])
            if field:
                key = field.group(1).lower()
                if key not in _FIELDS:
                    raise ValueError("Unsupported search field: " + key)
                result.append(("FIELD", key))
                pos += len(field.group())
            else:
                match = re.match(r'[^\s()"]+', query[pos:])
                text = match.group()
                upper = text.upper()
                result.append(
                    (upper if upper in {"AND", "OR", "ANDNOT"} else "TERM", text)
                )
                pos += len(text)
        if len(result) > 128:
            raise ValueError("Search query exceeds 128 tokens")
    return result


def compile_query(query):
    tokens = _tokens(query)
    if not tokens:
        return SearchQuery("TRUE", (), "0", (), None)
    pos = 0

    def primary(depth=0, inherited="all"):
        nonlocal pos
        if depth > 12 or pos >= len(tokens):
            raise ValueError("Missing search term or query nesting exceeds 12")
        kind, value = tokens[pos]
        pos += 1
        if kind == "FIELD":
            return primary(depth + 1, value)
        if kind == "(":
            node = expression(depth + 1, inherited)
            if pos >= len(tokens) or tokens[pos][0] != ")":
                raise ValueError("Unclosed search group")
            pos += 1
            return node
        if kind not in {"TERM", "PHRASE"}:
            raise ValueError("Expected search term")
        field = _FIELDS[inherited]
        vector = f"to_tsvector('english', {field})"
        if kind == "PHRASE":
            predicate = f"position(lower(%s) in lower({field})) > 0"
            rank = f"ts_rank({vector}, phraseto_tsquery('english', %s))"
        else:
            predicate = f"{vector} @@ plainto_tsquery('english', %s)"
            rank = f"ts_rank({vector}, plainto_tsquery('english', %s))"
        return predicate, (value,), rank, (value,)

    def conjunction(depth, inherited):
        nonlocal pos
        node = primary(depth, inherited)
        while pos < len(tokens) and tokens[pos][0] not in {"OR", ")"}:
            operator = tokens[pos][0]
            if operator in {"AND", "ANDNOT"}:
                pos += 1
            else:
                operator = "AND"
            right = primary(depth, inherited)
            pred = (
                f"({node[0]}) AND "
                + ("NOT " if operator == "ANDNOT" else "")
                + f"({right[0]})"
            )
            # Excluded terms must not raise a result's relevance rank.
            rank = node[2] if operator == "ANDNOT" else f"({node[2]}) + ({right[2]})"
            rp = node[3] if operator == "ANDNOT" else node[3] + right[3]
            node = pred, node[1] + right[1], rank, rp
        return node

    def expression(depth, inherited):
        nonlocal pos
        node = conjunction(depth, inherited)
        while pos < len(tokens) and tokens[pos][0] == "OR":
            pos += 1
            right = conjunction(depth, inherited)
            node = (
                f"({node[0]}) OR ({right[0]})",
                node[1] + right[1],
                f"({node[2]}) + ({right[2]})",
                node[3] + right[3],
            )
        return node

    result = expression(0, "all")
    if pos != len(tokens):
        raise ValueError("Unexpected search group terminator")
    plain = query.strip() if all(t[0] == "TERM" for t in tokens) else None
    return SearchQuery(*result, plain)
