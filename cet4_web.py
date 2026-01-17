import re
from collections import Counter
from pathlib import Path
from typing import Dict, List, Optional, Set

from flask import Flask, jsonify, render_template, request
from markupsafe import Markup, escape

WORD_FILE = Path(__file__).resolve().parent / "wordscheck" / "CET4_words_from_CET46_2016.csv"
TOKEN_RE = re.compile(r"[A-Za-z]+(?:[-'][A-Za-z]+)*")

TRANSLATION_TABLE = str.maketrans(
    {
        "\u2018": "'",
        "\u2019": "'",
        "\u2013": "-",
        "\u2014": "-",
        "\u2212": "-",
    }
)

app = Flask(__name__)


def normalize_text(text: str) -> str:
    return text.translate(TRANSLATION_TABLE)


def extract_tokens(text: str) -> List[str]:
    return TOKEN_RE.findall(normalize_text(text))


def normalize_token(token: str) -> str:
    return token.lower()


IRREGULAR_FORMS = {
    "am": "be",
    "is": "be",
    "are": "be",
    "was": "be",
    "were": "be",
    "been": "be",
    "being": "be",
    "does": "do",
    "did": "do",
    "done": "do",
    "has": "have",
    "had": "have",
}

CONTRACTION_EXCEPTIONS = {
    "won't": ["will"],
    "can't": ["can"],
    "shan't": ["shall"],
    "ain't": ["am", "is", "are", "has", "have"],
}


def expand_parentheses(word: str) -> List[str]:
    match = re.search(r"\(([^()]*)\)", word)
    if not match:
        return [word]
    start, end = match.span()
    inside = match.group(1)
    before = word[:start]
    after = word[end:]
    variants = []
    for tail in expand_parentheses(after):
        variants.append(before + inside + tail)
        variants.append(before + tail)
    return list({variant for variant in variants if variant})


def apply_suffix_variant(base: str, suffix: str) -> str:
    if not suffix:
        return base
    if base.endswith("ction") and suffix == "xion":
        return base[:-5] + suffix
    if len(suffix) <= len(base):
        return base[: -len(suffix)] + suffix
    return base + suffix


def expand_entry(entry: str) -> Set[str]:
    parts = entry.split("/")
    base_part = parts[0]
    base_variants = expand_parentheses(base_part)
    variants = set(base_variants)
    for part in parts[1:]:
        if part.startswith("G") and base_variants:
            suffix = part[1:]
            for base in base_variants:
                variants.add(apply_suffix_variant(base, suffix))
        else:
            for variant in expand_parentheses(part):
                variants.add(variant)
    return variants


def normalize_word(word: str) -> Optional[str]:
    word = word.strip()
    if not word:
        return None
    word = word.replace("G", "-")
    return word.lower()


def load_cet4_words(word_file: Path) -> Set[str]:
    words: Set[str] = set()
    with word_file.open(encoding="utf-8") as file:
        for line in file:
            entry = line.strip()
            if not entry or entry.lower() == "word":
                continue
            for variant in expand_entry(entry):
                normalized = normalize_word(variant)
                if not normalized:
                    continue
                words.add(normalized)
                if "-" in normalized:
                    words.add(normalized.replace("-", ""))
    return words


CET4_WORDS = load_cet4_words(WORD_FILE) if WORD_FILE.exists() else set()


def decode_uploaded_text(raw: bytes) -> str:
    for encoding in ("utf-8", "gbk"):
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="ignore")


def contraction_bases(token: str) -> List[str]:
    if token in CONTRACTION_EXCEPTIONS:
        return CONTRACTION_EXCEPTIONS[token][:]
    bases = []
    if token.endswith("n't") and len(token) > 3:
        bases.append(token[:-3])
    for suffix in ("'re", "'ve", "'ll", "'d", "'m"):
        if token.endswith(suffix):
            bases.append(token[: -len(suffix)])
    if token.endswith("'s"):
        bases.append(token[:-2])
    if token.endswith("'"):
        bases.append(token[:-1])
    return [base for base in bases if base]


def stem_ly(token: str) -> Set[str]:
    candidates: Set[str] = set()
    if token.endswith("ly") and len(token) > 4:
        base = token[:-2]
        candidates.add(base)
        if base.endswith("i") and len(base) > 1:
            candidates.add(base[:-1] + "y")
    return candidates


def stem_plural(token: str) -> Set[str]:
    candidates: Set[str] = set()
    if len(token) <= 3:
        return candidates
    if token.endswith("ies") and len(token) > 4:
        candidates.add(token[:-3] + "y")
    if token.endswith("ves") and len(token) > 4:
        candidates.add(token[:-3] + "f")
        candidates.add(token[:-3] + "fe")
    if token.endswith("es") and len(token) > 3:
        candidates.add(token[:-2])
    if token.endswith("s") and len(token) > 3:
        candidates.add(token[:-1])
    return candidates


def stem_ing(token: str) -> Set[str]:
    candidates: Set[str] = set()
    if token.endswith("ing") and len(token) > 5:
        base = token[:-3]
        candidates.add(base)
        if len(base) > 1 and base[-1] == base[-2]:
            candidates.add(base[:-1])
        if not base.endswith("e"):
            candidates.add(base + "e")
    return candidates


def stem_ed(token: str) -> Set[str]:
    candidates: Set[str] = set()
    if token.endswith("ed") and len(token) > 4:
        base = token[:-2]
        candidates.add(base)
        if len(base) > 1 and base[-1] == base[-2]:
            candidates.add(base[:-1])
        if base.endswith("i"):
            candidates.add(base[:-1] + "y")
        if not base.endswith("e"):
            candidates.add(base + "e")
    return candidates


def stem_er(token: str) -> Set[str]:
    candidates: Set[str] = set()
    if token.endswith("er") and len(token) > 4:
        base = token[:-2]
        candidates.add(base)
        if base.endswith("i") and len(base) > 1:
            candidates.add(base[:-1] + "y")
        if len(base) > 1 and base[-1] == base[-2]:
            candidates.add(base[:-1])
        if not base.endswith("e"):
            candidates.add(base + "e")
    return candidates


def stem_est(token: str) -> Set[str]:
    candidates: Set[str] = set()
    if token.endswith("est") and len(token) > 5:
        base = token[:-3]
        candidates.add(base)
        if base.endswith("i") and len(base) > 1:
            candidates.add(base[:-1] + "y")
        if len(base) > 1 and base[-1] == base[-2]:
            candidates.add(base[:-1])
        if not base.endswith("e"):
            candidates.add(base + "e")
    return candidates


def generate_candidates(token: str) -> Set[str]:
    candidates: Set[str] = {token}
    if "-" in token:
        candidates.add(token.replace("-", ""))
    if token in IRREGULAR_FORMS:
        candidates.add(IRREGULAR_FORMS[token])
    candidates.update(contraction_bases(token))

    for base in list(candidates):
        if base in IRREGULAR_FORMS:
            candidates.add(IRREGULAR_FORMS[base])
        candidates.update(stem_ly(base))
        candidates.update(stem_plural(base))
        candidates.update(stem_ing(base))
        candidates.update(stem_ed(base))
        candidates.update(stem_er(base))
        candidates.update(stem_est(base))

    return {candidate for candidate in candidates if candidate}


def is_known_word(token: str) -> bool:
    normalized = normalize_token(token)
    for candidate in generate_candidates(normalized):
        if candidate in CET4_WORDS:
            return True
    return False


def highlight_missing(text: str) -> Markup:
    normalized = normalize_text(text)
    output: List[str] = []
    last_index = 0

    for match in TOKEN_RE.finditer(normalized):
        start, end = match.span()
        original_chunk = text[start:end]
        output.append(str(escape(text[last_index:start])))
        if is_known_word(normalized[start:end]):
            output.append(str(escape(original_chunk)))
        else:
            output.append(f'<span class="missing">{escape(original_chunk)}</span>')
        last_index = end

    output.append(str(escape(text[last_index:])))
    return Markup("".join(output))


def build_results(text: str) -> Dict[str, object]:
    tokens = extract_tokens(text)
    missing = []
    for token in tokens:
        if not is_known_word(token):
            missing.append(normalize_token(token))
    counter = Counter(missing)
    missing_items = sorted(counter.items(), key=lambda item: (-item[1], item[0]))
    return {
        "total_count": len(tokens),
        "missing_count": sum(counter.values()),
        "unique_missing": len(counter),
        "missing_items": missing_items,
        "unique_list": "\n".join(word for word, _ in missing_items),
        "highlighted": str(highlight_missing(text)),
    }


@app.route("/", methods=["GET", "POST"])
def index() -> str:
    input_text = ""
    error: Optional[str] = None
    results: Optional[Dict[str, object]] = None

    if request.method == "POST":
        file_storage = request.files.get("text_file")
        if file_storage and file_storage.filename:
            raw = file_storage.read()
            input_text = decode_uploaded_text(raw)
        else:
            input_text = request.form.get("text_input", "")

        if not input_text.strip():
            error = "Please upload a text file or paste some text."
        elif not CET4_WORDS:
            error = "CET4 word list not found."
        else:
            results = build_results(input_text)

    return render_template(
        "index.html",
        input_text=input_text,
        error=error,
        results=results,
    )


@app.route("/check", methods=["POST"])
def check() -> tuple[str, int]:
    if not CET4_WORDS:
        return jsonify({"error": "CET4 word list not found."}), 500

    data = request.get_json(silent=True) or {}
    text = data.get("text", "")
    if not isinstance(text, str):
        text = ""

    if not text.strip():
        return (
            jsonify(
                {
                    "total_count": 0,
                    "missing_count": 0,
                    "unique_missing": 0,
                    "missing_items": [],
                    "unique_list": "",
                    "highlighted": str(escape(text)),
                }
            ),
            200,
        )

    results = build_results(text)
    results["missing_items"] = [
        {"word": word, "count": count} for word, count in results["missing_items"]
    ]
    return jsonify(results), 200


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=False)
