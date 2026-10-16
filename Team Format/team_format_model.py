"""Team format model for SWWU: builds per-leg team scores, league tables and participation tables
for each combination of format proposal (unchanged / P1 / P1a) and gender requirement (P2 off/on)."""
from itertools import combinations
from pathlib import Path

import pandas as pd
import archeryutils as au
from archeryutils import handicaps as hc

DATA = Path(__file__).resolve().parent.parent / "Clean Data" / "completeResults.csv"
WA = au.load_rounds.WA_outdoor
ROUNDS = {"WA 70m": WA.wa720_70, "WA 60m": WA.wa720_60, "WA 50m (122cm)": WA.wa720_50_b, "WA 50m (80cm)": WA.wa720_50_c}
LEGS = list(range(1, 9))
POINTS = {r: 10 - (r - 1) for r in range(1, 9)}  # 10 for 1st, 9 for 2nd, ...
NONCOMP, UNSIGHTED = {"R", "B", "L"}, {"B", "L"}


def load_archers() -> pd.DataFrame:
    """Load results; drop alumni and non-starters (score 0).

    Returns: DataFrame with columns club, leg, div, novice (bool), gender ('F'/'O'), score, round (shot round name),
    tb_order (e.g. 'Hits>Golds'), tb (dict of tiebreak counts, 0 where n/a), name.
    """
    d = pd.read_csv(DATA)
    d = d[(d["Score"] > 0) & (~d["Class"].str.startswith("A"))]
    return pd.DataFrame({
        "club": d["Country"], "leg": d["Leg"], "div": d["Division"], "novice": d["Class"].str.startswith("N"),
        "gender": d["Gender"], "score": d["Score"], "round": d["Round Shot"],
        "tb_order": d["TB Order"],
        "tb": d[["TB Hits", "TB Golds", "TB Xs", "TB 10s", "TB 9s"]].fillna(0).to_dict("records"),
        "name": d["GivenName"] + " " + d["FamilyName"]}).reset_index(drop=True)


def convert(score: int, shot: str, target: str) -> int:
    """Convert a score on round `shot` to the handicap-equivalent score on round `target` (Archery GB scheme).

    Args: score (int), shot/target (round names keyed in ROUNDS). Returns: int equivalent score.
    """
    if shot == target:
        return int(score)
    handicap = hc.handicap_from_score(score, ROUNDS[shot], "AGB")
    return int(hc.score_for_round(handicap, ROUNDS[target], "AGB"))


# Team definitions: name, size, eligibility predicate, leg-8 target round (None = no conversion), max compound archers
def formats(proposal: str) -> list[dict]:
    """Return team definitions for 'unchanged', 'P1' or 'P1a'.

    Each dict: name (str), n (int team size), ok (callable on archer row -> bool), target (round name or None),
    max_c (max compound archers or None), legs (list of legs that count).
    """
    novice_nc = lambda a: a.div in NONCOMP and a.novice
    if proposal == "unchanged":
        return [dict(name="ENCT", n=4, ok=lambda a: a.div in NONCOMP, target="WA 70m", max_c=None, legs=LEGS),
                dict(name="NNCT", n=3, ok=novice_nc, target="WA 60m", max_c=None, legs=LEGS),
                dict(name="CT", n=2, ok=lambda a: a.div == "C", target=None, max_c=None, legs=LEGS)]
    ust = dict(name="UST", n=3, ok=lambda a: a.div in UNSIGHTED, target="WA 50m (122cm)", max_c=None)
    nnct = dict(name="NNCT", n=3, ok=novice_nc, target="WA 60m", max_c=None)
    if proposal == "P1":
        st = dict(name="ENCT", n=3, ok=lambda a: a.div in NONCOMP, target="WA 70m", max_c=None)
        return [{**t, "legs": LEGS} for t in (st, nnct, ust)]
    st = dict(name="ST", n=3, ok=lambda a: a.div in {"R", "C"}, target="WA 70m", max_c=1)
    return [{**t, "legs": [1, 2, 3, 4, 5, 6]} for t in (st, nnct, ust)]


def select_team(pool: list[tuple], n: int, max_c, gender: bool):
    """Pick the scoring archers for one club/leg/team. Incomplete teams are allowed (any pool of >=1 archer scores).

    The team has up to n archers (best total score). If gender is True: when the pool has both an F and an O the team
    must contain >=1 of each; when it lacks either, at most n-1 archers contribute (a gender shortfall).

    Args: pool list of (score, gender, is_compound, name, tb_order, tb); n team size; max_c max compound archers
    (None = no limit); gender bool, whether the gender requirement applies.
    Returns: (chosen list of tuples, lacked_gender bool) or (None, False) if the pool is empty.
    """
    def best_subset(size, need_both):
        for k in range(min(size, len(pool)), 0, -1):
            best = None
            for combo in combinations(pool, k):
                if max_c is not None and sum(p[2] for p in combo) > max_c:
                    continue
                if need_both and not ({"F", "O"} <= {p[1] for p in combo}):
                    continue
                total = sum(p[0] for p in combo)
                if best is None or total > best[0]:
                    best = (total, combo)
            if best:
                return list(best[1])
        return None

    if not pool:
        return None, False
    lacked = gender and not ({"F", "O"} <= {p[1] for p in pool})
    if lacked:
        return best_subset(n - 1, False), len(pool) >= n
    return best_subset(n, gender), False


TB_COL = {"Hits": "TB Hits", "Golds": "TB Golds", "Xs": "TB Xs", "10s": "TB 10s", "9s": "TB 9s"}


def team_tiebreak(chosen: list[tuple]) -> tuple:
    """Team tiebreak key: member counts summed, in the order given by the round's ruleset (e.g. Hits>Golds).

    Args: chosen list of pool tuples. Returns: tuple of summed counts (higher is better, compared in order).
    """
    return tuple(sum(p[5][TB_COL[k]] for p in chosen) for k in chosen[0][4].split(">"))


def compute(arch: pd.DataFrame, proposal: str, gender: bool) -> dict:
    """Compute results for one combination.

    Returns dict with 'teams' {team name: dict(score DataFrame club x leg, count DataFrame, rank DataFrame,
    lacked DataFrame (bool), table DataFrame league table)}.
    """
    clubs = sorted(arch["club"].unique())
    out = {}
    for t in formats(proposal):
        score = pd.DataFrame(index=clubs, columns=LEGS, dtype=float)
        count = pd.DataFrame(0, index=clubs, columns=LEGS)
        lacked = pd.DataFrame(False, index=clubs, columns=LEGS)
        members = {}
        tie = pd.DataFrame(None, index=clubs, columns=LEGS, dtype=object)
        for leg in t["legs"]:
            la = arch[arch["leg"] == leg]
            for club in clubs:
                pool = []
                for a in la[la["club"] == club].itertuples():
                    if not t["ok"](a):
                        continue
                    s = convert(a.score, a.round, t["target"]) if leg == 8 and t["target"] else a.score
                    pool.append((s, a.gender, a.div == "C", a.name, a.tb_order, a.tb))
                chosen, lk = select_team(pool, t["n"], t["max_c"], gender)
                lacked.loc[club, leg] = lk
                if chosen:
                    score.loc[club, leg] = sum(p[0] for p in chosen)
                    tie.loc[club, leg] = team_tiebreak(chosen)
                    count.loc[club, leg] = len(chosen)
                    members[(club, leg)] = chosen
        rank = pd.DataFrame(index=clubs, columns=LEGS, dtype=float)
        tied = []
        for leg in LEGS:
            keys = {c: (score.loc[c, leg],) + tie.loc[c, leg] for c in clubs if pd.notna(score.loc[c, leg])}
            for c, k in keys.items():
                rank.loc[c, leg] = 1 + sum(o > k for o in keys.values())
            tied += [(c, leg) for c, k in keys.items() if sum(o == k for o in keys.values()) > 1]
        pts = rank.apply(lambda c: c.map(POINTS)).fillna(0)
        table = pd.DataFrame({"LP": pts.sum(axis=1), "Agg": score.sum(axis=1)})
        table = table.sort_values(["LP", "Agg"], ascending=False)
        table["Pos"] = range(1, len(table) + 1)
        out[t["name"]] = dict(score=score, count=count, rank=rank, lacked=lacked, table=table, members=members, unresolved_ties=tied,
                              legs=t["legs"], n=t["n"])
    return out


def run_all() -> dict:
    """Return {(proposal, gender): compute(...)} for all six combinations."""
    arch = load_archers()
    return {(p, g): compute(arch, p, g) for p in ("unchanged", "P1", "P1a") for g in (False, True)}


if __name__ == "__main__":
    res = run_all()
    for k, v in res.items():
        for tn, r in v.items():
            print(k, tn)
            print(r["table"].join(r["rank"]).to_string())
