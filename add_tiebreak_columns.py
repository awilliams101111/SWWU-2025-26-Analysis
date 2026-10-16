"""Append tiebreak columns to Clean Data/completeResults.csv using the raw per-arrow data in Raw Data/.

New columns (appended, existing columns untouched; no archers added or removed):
  TB Order  - tiebreak ranking for the round shot, e.g. 'Hits>Golds' (compared in that order, higher is better)
  TB Hits, TB Golds, TB Xs, TB 10s, TB 9s - counts for the archer's whole round (blank where not applicable)

Golds: Portsmouth/Bray = 10s; Worcester = 5s; WA 900 / WA 720 = 10s + Xs. Leg 1 comes from the Leg 1 PDF
(already in the clean data as Hits and '10' = Golds).
"""
import csv
from pathlib import Path

ROOT = Path(__file__).resolve().parent
CLEAN = ROOT / "Clean Data" / "completeResults.csv"
RAW = ROOT / "Raw Data"
NEW_COLS = ["TB Order", "TB Hits", "TB Golds", "TB Xs", "TB 10s", "TB 9s"]
# leg -> (raw file, tiebreak order, golds definition)
LEGS = {2: ("SWWU25L2.txt", "10s>9s"), 3: ("SWWU25L3.txt", "Hits>Golds"), 4: ("SWWU26L4.txt", "Hits>Golds>Xs"),
        5: ("SWWU26L5.txt", "Hits>Golds"), 6: ("SWWU26L6.txt", "Hits>Golds"), 7: ("SWWU26L7.txt", "Golds>Xs"),
        8: ("SWWU26L8.txt", "Golds>Xs")}


def count_arrows(raw_row: list[str], header: list[str]) -> dict:
    """Count arrow values over all 'Arrows' fields of a raw row.

    Args: raw_row (list of str fields), header (list of str). Returns: dict of counts: hits, x, n10, n9, n5.
    """
    arrows = [a for h, a in zip(header, raw_row) if h.endswith("Arrows") and a]
    vals = [v.strip() for s in arrows for v in s.split(",") if v.strip()]
    return dict(hits=sum(v != "M" for v in vals), x=vals.count("X"), n10=vals.count("10"), n9=vals.count("9"),
                n5=vals.count("5"))


def reported_counts(leg: int, row: list[str], header: list[str]) -> dict:
    """Counts from a raw row's own summary columns (used when no arrow-by-arrow data was entered).

    Args: leg (int), row (list of str), header (list of str). Returns: dict of counts as in count_arrows. Hits are the
    file's 'Hits' value (arrows shot), the best available where arrows are missing.
    """
    i = header.index("Score")
    hits, b, c = (int(row[i + k] or 0) for k in (1, 2, 3))
    if leg in (3, 5, 6):  # duplicate 'Hits' columns: the first can be 0, the second holds the arrow count
        hits = max(hits, b)
    z = dict(hits=hits, x=0, n10=0, n9=0, n5=0)
    if leg == 2:
        z.update(n10=b, n9=c)
    elif leg == 4:
        z.update(n5=b, x=c)
    elif leg == 7:
        z.update(n10=b - c, x=c)  # raw '10+X' minus X
    elif leg == 8:
        z.update(x=b, n10=c)
    else:
        z.update(n10=c)  # Golds column is the last of the summary block
    return z


def tiebreak_fields(leg: int, order: str, c: dict) -> list:
    """Return [Order, Hits, Golds, Xs, 10s, 9s] for a leg from arrow counts c."""
    hits, xs, tens, nines, fives = c["hits"], c["x"], c["n10"], c["n9"], c["n5"]
    if leg == 2:
        return [order, hits, "", "", tens, nines]
    if leg == 4:
        return [order, hits, fives, xs, "", ""]
    if leg in (7, 8):
        return [order, hits, tens + xs, xs, "", ""]
    return [order, hits, tens, "", "", ""]


def main():
    with open(CLEAN, newline="", encoding="utf-8-sig") as f:
        rows = list(csv.reader(f))
    head, body = rows[0], rows[1:]
    if "TB Order" in head:  # idempotent: drop previously added columns
        n = head.index("TB Order")
        head, body = head[:n], [r[:n] for r in body]
    ix = {n: head.index(n) for n in ("FamilyName", "GivenName", "Score", "Leg", "Hits", "10")}
    # raw lookup: (leg, family, given, score) -> tiebreak fields
    lookup, bad = {}, []
    for leg, (fname, order) in LEGS.items():
        with open(RAW / fname, newline="", encoding="utf-8-sig") as f:
            r = list(csv.reader(f, delimiter=";"))
        h = r[0]
        for row in r[1:]:
            if len(row) < 25:
                continue
            c = count_arrows(row, h)
            if c["hits"] == 0:
                c = reported_counts(leg, row, h)
            key = (leg, row[2].strip().lower(), row[3].strip().lower(), int(row[h.index("Score")]))
            lookup.setdefault(key, []).append(tiebreak_fields(leg, order, c))
            # validate against the file's own summary columns (arrow totals vs reported hits / tens etc.)
            rep_hits = int(row[h.index("Score") + 1] or 0)
            if rep_hits != c["hits"]:
                bad.append((leg, key, "hits", rep_hits, c["hits"]))
    out, miss = [], []
    for row in body:
        leg, score = int(row[ix["Leg"]]), int(row[ix["Score"]])
        if leg == 1:
            new = ["Hits>Golds", row[ix["Hits"]], row[ix["10"]], "", "", ""] if score > 0 else [""] * 6
        else:
            key = (leg, row[ix["FamilyName"]].strip().lower(), row[ix["GivenName"]].strip().lower(), score)
            found = lookup.get(key)
            new = found[0] if found else [LEGS[leg][1]] + [""] * 5
            if not found and score > 0:
                miss.append(key)
        out.append(row + new)
    with open(CLEAN, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(head + NEW_COLS)
        w.writerows(out)
    print("rows", len(out), "unmatched scoring rows:", miss, "hit mismatches:", bad[:5])


if __name__ == "__main__":
    main()
