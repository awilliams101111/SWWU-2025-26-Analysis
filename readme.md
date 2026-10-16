# SWWU 2025-26 Analysis

Analysis of scores of archers Participating in the SWWU Archery League in 2025-26

## Status

### Done: team format analysis
Re-scores the 2025-26 league under the proposed team formats, with and without the gender requirement (6 combinations: Unchanged / Proposal 1 / Proposal 1a, each with Proposal 2 off and on). Proposals are defined in `Team Format/TeamFormatPropsals.md`.

1. Effect of requiring both gender categories to be represented in a team (Proposal 2) - done
2. Effect of changing ENCT format to ST of 3 archers, max 1 compound, indoor legs only (Proposal 1a) - done
3. Effect of changing ENCT to 3 archers, no compounds (Proposal 1; barebow/longbow/recurve eligible) - done. A recurve-only team (RT) was not analysed.
4. Effect of adding UST (Proposals 1 and 1a) - done

Output: `Team Format/teamFormatAnalysis.xlsx` (shareable on its own)
- `Explanation` sheet: formats, rules and assumptions.
- `Team format analysis` sheet: per-leg league tables, participation tables (colour scaled) and an impact comment for each table; rows are the formats, columns are Proposal 2 off/on.

Headline findings (details in the workbook):
- Barebow/longbow are ~40% of archer-legs and compound only ~5.5%. The UST fielded a team in 56 of 64 club-legs; the compound team only 26 (full only 5).
- Proposal 2 has a small effect on results: 2-5 shortfall cases per format, each costing a team 12-33% of its leg score. It changes the overall league order only in the Proposal 1a ST. Female attendance is lower outdoors (about 24-29% in legs 7-8 vs about 40% indoors), so it bites hardest there.
- Proposal 1a drops legs 7-8, the least attended legs, and removes barebow/longbow from the ST.
- Suggested option: Proposal 1 structure with Proposal 2 and all 8 legs. Proposal 1 excludes compound entirely, so a variant (a single compound slot in the ENCT in every leg, needing the handicap conversion applied to leg 7 as well as leg 8) is worth testing. Proposal 2 applied to indoor legs only is also untested.

### Outstanding
5. Find handicap/classification of each score shot and plot handicaps per leg for each archer over the year - not started. `archeryutils` (already a dependency) provides the handicap and classification functions.

## Files
- `Clean Data/completeResults.csv` - consolidated results. Tiebreak columns `TB Order`, `TB Hits`, `TB Golds`, `TB Xs`, `TB 10s`, `TB 9s` were appended by `add_tiebreak_columns.py` from `Raw Data/` (the raw `Hits` column is often arrows shot, not hits, so use `TB Hits`).
- `add_tiebreak_columns.py` - rebuilds the tiebreak columns (safe to re-run).
- `Team Format/team_format_model.py` - team selection, handicap conversion, league tables.
- `Team Format/build_excel.py` - writes the workbook. It rebuilds the file from scratch, so manual edits to the workbook are overwritten.
- `participation.ipynb` - participation and medal counts.

Re-run the workbook: `uv run python "Team Format/build_excel.py"`

## Key modelling assumptions
- Alumni excluded; archers may score in more than one team; incomplete teams are allowed.
- Proposal 2: best team containing at least one Female and one Open archer (all combinations tested); if the eligible pool lacks either, only team size - 1 archers count.
- League points 10, 9, 8...; league ties broken on aggregate; equal team scores broken by the round's rule summed over the team (see the Explanation sheet).
- Leg 8 scores converted with the Archery GB handicap system (`archeryutils`) to WA 70m (ENCT/ST), WA 60m (NNCT) or WA 50m 122cm (UST). The CT is not converted.
- Gender as recorded per archer per leg.

# Acronyms
- LX: Leg X
- ENCT: Experienced Non-Compound Team
- NNCT: Novice Non-Compound Team
- CT: Compound Team
- ST: Sighted Team (Proposal 1a only)
- UST: Unsighted Team
- RT: Recurve Team (not analysed)
- LP: League Points
- Agg: Aggregate (sum of team scores)
