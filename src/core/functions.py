import re


def clean_refs(text: str) -> str:
    if not text:
        return text
    text = re.sub(r"\[\s*\d+\s*\]", "", text)
    text = re.sub(r"\(\s*\d+\s*\)", "", text)
    text = re.sub(r"\s{2,}", " ", text)
    text = re.sub(r"\s+([,.:!?])", r"\1", text)
    return text.strip()


def truncate_words(text: str, max_words: int = 120) -> str:
    if not text:
        return text
    w = text.split()
    return " ".join(w[:max_words]) + ("…" if len(w) > max_words else "")
