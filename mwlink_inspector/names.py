"""Normalisasi & fuzzy-matching nama stasiun (nama di screenshot sering tidak persis sama dengan ISR)."""
import re
try:
    from rapidfuzz import fuzz
except ImportError:
    import difflib
    class DummyFuzz:
        @staticmethod
        def ratio(s1, s2):
            return difflib.SequenceMatcher(None, s1, s2).ratio() * 100
        @staticmethod
        def token_set_ratio(s1, s2):
            t1 = " ".join(sorted(set(s1.split())))
            t2 = " ".join(sorted(set(s2.split())))
            return difflib.SequenceMatcher(None, t1, t2).ratio() * 100
    fuzz = DummyFuzz()



def norm(s: str) -> str:
    s = s.upper() if isinstance(s, str) else ""
    s = re.sub(r"\(.*?\)", " ", s)          # buang "(TELKOM)"
    s = re.sub(r"^\s*TO\s+", "", s)         # buang awalan "TO "
    s = re.sub(r"^\d+\s*/\s*", "", s)       # buang prefix ISR "2/BPP047_..."
    return re.sub(r"[^A-Z0-9]+", " ", s).strip()


def name_score(a: str, b: str) -> float:
    na, nb = norm(a), norm(b)
    if not na or not nb:
        return 0.0
    ca, cb = na.replace(" ", ""), nb.replace(" ", "")
    score = max(fuzz.token_set_ratio(na, nb), fuzz.ratio(ca, cb))
    short, long_ = sorted((ca, cb), key=len)
    if len(short) >= 8 and short in long_:   # satu nama berisi nama lain (mis. akhiran "_1K")
        score = max(score, 95)
    return float(score)
