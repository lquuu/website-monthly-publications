# website_monthly_pubs.py
#
# Builds two Excel documents tracking BI faculty publications, month by month,
# for a specified year. Each run is saved under its own timestamped filename --
# it never overwrites a previous run's output, so any manual edits you've made
# to an earlier file are safe:
#
#   website_monthly_pubs_<year>_<timestamp>.xlsx
#       (Title, Date, Journal, Contributors, Link, Clusters, Facilities, Publish Date)
#   website_monthly_pubs_import_template_<year>_<timestamp>.xlsx
#       (Title, Date, Journal, Contributors, Link to Publication, Clusters, Publish Date)
#
# THERE ARE TWO WAYS TO SPECIFY THE YEAR (default is the current year):
# 1) Command-line argument: python website_monthly_pubs.py 2026
# 2) Hard code: TARGET_YEAR = 2025 below
#
# *** IMPORTANT CHANGE ***
# This version is fully INDEPENDENT of bi_faculty_citations_by_publication.csv and
# of pyblio_draft4_BI.py / bi_institutional_metrics.py. It no longer needs either of
# those to have run first. Instead, it queries Scopus directly, scoped to just the
# target year, via ScopusSearch -- which is dramatically cheaper than the original
# collection script because:
#   - pyblio_draft4_BI.py calls get_document_eids(refresh=True) per faculty member,
#     which forces a fresh fetch of that person's ENTIRE career history every run,
#     then calls CitationOverview per publication on top of that.
#   - This script instead asks Scopus directly for "papers by these faculty IDs,
#     published in year Y" -- a handful of batched queries total, each one cached
#     locally for 30 days (see refresh=30 below), and each query already returns
#     title/journal/date/DOI/full author list in one go (no AbstractRetrieval needed
#     at all).
# Since results are cached for 30 days, running this monthly will mostly hit the
# API fresh (since each run is ~30 days after the last), which is exactly the
# cadence you described -- but running it twice in the same week, e.g. while
# testing, will hit the local cache and cost nothing.
#
# NOTES:
# - Contributors and Clusters are derived from each paper's actual author byline
#   as returned directly by ScopusSearch (author_names / author_ids fields),
#   filtered down to whichever of your bi_faculty IDs appear on that paper, in
#   the paper's own author order.
# - Faculty names, cluster memberships, and all known primary/secondary Scopus IDs
#   are maintained together in BI_FACULTY_ROSTER below.
# - Publications are matched to faculty directly by Scopus ID, so duplicate surnames,
#   punctuation, initials, and alternate Scopus profiles do not require special cases.
# - Publications with no DOI are skipped (can't build a Link), same as the original
#   collection script's behavior.

import sys
import calendar
from datetime import date, datetime

import pandas as pd
from pybliometrics.scopus import init as scopus_init, ScopusSearch

scopus_init()

# ---- Year to run for -----------------------------------------------------
TARGET_YEAR = int(sys.argv[1]) if len(sys.argv) > 1 else datetime.now().year
RUN_TIMESTAMP = datetime.now().strftime("%Y%m%d_%H%M%S")
print(f"Building monthly pub tracker for {TARGET_YEAR}...")

# Each run gets its own timestamped files -- never overwrites a prior run's output,
# so any manual edits you've made to a previous file are untouched.
OUT_MAIN = f"website_monthly_pubs_{TARGET_YEAR}_{RUN_TIMESTAMP}.xlsx"
OUT_TEMPLATE = f"website_monthly_pubs_import_template_{TARGET_YEAR}_{RUN_TIMESTAMP}.xlsx"

# ---- BI faculty roster ----------------------------------------------------
# ============================================================================
# BI FACULTY ROSTER
#
# This is the only section that should need updating when faculty join or leave.
# Each faculty record contains separate first, middle, and last name fields, the
# contributor format used in the spreadsheet, every known Scopus Author ID, and
# BI cluster memberships. Include ALL IDs when Scopus has duplicate profiles.
# ============================================================================
BI_FACULTY_ROSTER = [
    {
        "first": 'Carlos',
        "middle": 'A.',
        "last": 'Aguilar',
        "contributor": 'Aguilar, C. A.',
        "scopus_ids": [55016188700],
        "clusters": ['Cell and Tissue Engineering', 'Single Cell Technologies'],
    },
    {
        "first": 'Timothy',
        "middle": 'M.',
        "last": 'Bruns',
        "contributor": 'Bruns, T. M.',
        "scopus_ids": [24068798200],
        "clusters": ['Neural Engineering'],
    },
    {
        "first": 'Cynthia',
        "middle": 'A.',
        "last": 'Chestek',
        "contributor": 'Chestek, C. A.',
        "scopus_ids": [12779584500],
        "clusters": ['Neural Engineering'],
    },
    {
        "first": 'Tae-Hwa',
        "middle": '',
        "last": 'Chun',
        "contributor": 'Chun, T. H.',
        "scopus_ids": [57195268661],
        "clusters": ['Cell and Tissue Engineering'],
    },
    {
        "first": 'Lola',
        "middle": '',
        "last": 'Eniola-Adefeso',
        "contributor": 'Eniola-Adefeso, L.',
        "scopus_ids": [26021803700],
        "clusters": ['Cell and Tissue Engineering'],
    },
    {
        "first": 'Eva',
        "middle": 'L.',
        "last": 'Feldman',
        "contributor": 'Feldman, E. L.',
        "scopus_ids": [7201917779],
        "clusters": [],
    },
    {
        "first": 'Sharon',
        "middle": 'C.',
        "last": 'Glotzer',
        "contributor": 'Glotzer, S. C.',
        "scopus_ids": [7004361447],
        "clusters": ['Nanotechnology'],
    },
    {
        "first": 'Colin',
        "middle": 'F.',
        "last": 'Greineder',
        "contributor": 'Greineder, C. F.',
        "scopus_ids": [15033338100],
        "clusters": ['Nanotechnology'],
    },
    {
        "first": 'Toshiro',
        "middle": '',
        "last": 'Hara',
        "contributor": 'Hara, T.',
        "scopus_ids": [55465936200],
        "clusters": ['BioInnovations in Brain Cancer'],
    },
    {
        "first": 'Jan',
        "middle": '',
        "last": 'Hu',
        "contributor": 'Hu, J.',
        "scopus_ids": [8072516000, 57201972955],
        "clusters": ['Cell and Tissue Engineering'],
    },
    {
        "first": 'Evan',
        "middle": 'T.',
        "last": 'Keller',
        "contributor": 'Keller, E. T.',
        "scopus_ids": [7202812560],
        "clusters": ['Cell and Tissue Engineering', 'Single Cell Technologies'],
    },
    {
        "first": 'Jinsang',
        "middle": '',
        "last": 'Kim',
        "contributor": 'Kim, J.',
        "scopus_ids": [9536007500, 57203325389],
        "clusters": ['Advanced Materials and Drug Delivery'],
    },
    {
        "first": 'Nicholas',
        "middle": '',
        "last": 'Kotov',
        "contributor": 'Kotov, N.',
        "scopus_ids": [22135180200],
        "clusters": ['Nanotechnology', 'Single Cell Technologies'],
    },
    {
        "first": 'Joerg',
        "middle": '',
        "last": 'Lahann',
        "contributor": 'Lahann, J.',
        "scopus_ids": [6603555265],
        "clusters": ['Advanced Materials and Drug Delivery', 'BioInnovations in Brain Cancer', 'Single Cell Technologies'],
    },
    {
        "first": 'Ronald',
        "middle": 'G.',
        "last": 'Larson',
        "contributor": 'Larson, R. G.',
        "scopus_ids": [7402161971],
        "clusters": ['Nanotechnology'],
    },
    {
        "first": 'Scott',
        "middle": 'F.',
        "last": 'Lempka',
        "contributor": 'Lempka, S. F.',
        "scopus_ids": [16042824100, 57200610670],
        "clusters": ['Neural Engineering'],
    },
    {
        "first": 'Sasha Cai',
        "middle": '',
        "last": 'Lesher-Pérez',
        "contributor": 'Lesher-Pérez, S. C.',
        "scopus_ids": [36930755200],
        "clusters": ['Cell and Tissue Engineering', 'Single Cell Technologies'],
    },
    {
        "first": 'Albert',
        "middle": '',
        "last": 'Liu',
        "contributor": 'Liu, A.',
        "scopus_ids": [57189033142],
        "clusters": ['Advanced Materials and Drug Delivery', 'Nanotechnology', 'Cell and Tissue Engineering'],
    },
    {
        "first": 'Isabelle',
        "middle": 'M. A.',
        "last": 'Lombaert',
        "contributor": 'Lombaert, I. M. A.',
        "scopus_ids": [15035854800],
        "clusters": ['Cell and Tissue Engineering'],
    },
    {
        "first": 'Kathryn',
        "middle": '',
        "last": 'Luker',
        "contributor": 'Luker, K.',
        "scopus_ids": [7006008866],
        "clusters": ['Cell and Tissue Engineering', 'Single Cell Technologies'],
    },
    {
        "first": 'Jouha',
        "middle": '',
        "last": 'Min',
        "contributor": 'Min, J.',
        "scopus_ids": [55975937900],
        "clusters": ['Cell and Tissue Engineering', 'BioInnovations in Brain Cancer', 'Advanced Materials and Drug Delivery', 'Nanotechnology'],
    },
    {
        "first": 'James',
        "middle": 'J.',
        "last": 'Moon',
        "contributor": 'Moon, J. J.',
        "scopus_ids": [35726004100],
        "clusters": ['BioInnovations in Brain Cancer', 'Advanced Materials and Drug Delivery'],
    },
    {
        "first": 'Deepak',
        "middle": '',
        "last": 'Nagrath',
        "contributor": 'Nagrath, D.',
        "scopus_ids": [8577389100],
        "clusters": ['BioInnovations in Brain Cancer', 'Cell and Tissue Engineering', 'Single Cell Technologies'],
    },
    {
        "first": 'Sunitha',
        "middle": '',
        "last": 'Nagrath',
        "contributor": 'Nagrath, S.',
        "scopus_ids": [23101128900],
        "clusters": ['Cell and Tissue Engineering', 'Single Cell Technologies'],
    },
    {
        "first": 'Enrico',
        "middle": '',
        "last": 'Opri',
        "contributor": 'Opri, E.',
        "scopus_ids": [57189212345],
        "clusters": ['Neural Engineering'],
    },
    {
        "first": 'Abdon',
        "middle": '',
        "last": 'Pena-Francesch',
        "contributor": 'Pena-Francesch, A.',
        "scopus_ids": [55885015000],
        "clusters": ['Advanced Materials and Drug Delivery', 'Nanotechnology'],
    },
    {
        "first": 'Anthony',
        "middle": '',
        "last": 'Rosenzweig',
        "contributor": 'Rosenzweig, A.',
        "scopus_ids": [35243347700],
        "clusters": ['Cell and Tissue Engineering', 'Single Cell Technologies'],
    },
    {
        "first": 'Anna',
        "middle": 'A.',
        "last": 'Schwendeman',
        "contributor": 'Schwendeman, A. A.',
        "scopus_ids": [6506939515, 57222518153],
        "clusters": ['Nanotechnology', 'BioInnovations in Brain Cancer'],
    },
    {
        "first": 'Steven',
        "middle": 'P.',
        "last": 'Schwendeman',
        "contributor": 'Schwendeman, S. P.',
        "scopus_ids": [7003309460],
        "clusters": ['BioInnovations in Brain Cancer', 'Advanced Materials and Drug Delivery'],
    },
    {
        "first": 'Lonnie',
        "middle": 'D.',
        "last": 'Shea',
        "contributor": 'Shea, L. D.',
        "scopus_ids": [7005328686],
        "clusters": ['Cell and Tissue Engineering'],
    },
    {
        "first": 'Jae-Won',
        "middle": '',
        "last": 'Shin',
        "contributor": 'Shin, J.',
        "scopus_ids": [48061385600, 57205414505],
        "clusters": ['Nanotechnology', 'Cell and Tissue Engineering', 'Single Cell Technologies', 'Advanced Materials and Drug Delivery'],
        },
    {
        "first": 'Michael',
        "middle": 'J.',
        "last": 'Solomon',
        "contributor": 'Solomon, M. J.',
        "scopus_ids": [35307875200],
        "clusters": ['Nanotechnology'],
    },
    {
        "first": 'William',
        "middle": 'C.',
        "last": 'Stacey',
        "contributor": 'Stacey, W. C.',
        "scopus_ids": [7005915503],
        "clusters": ['Neural Engineering'],
    },
    {
        "first": 'Peter',
        "middle": 'M.',
        "last": 'Tessier',
        "contributor": 'Tessier, P. M.',
        "scopus_ids": [7101689432],
        "clusters": ['BioInnovations in Brain Cancer', 'Advanced Materials and Drug Delivery', 'Nanotechnology'],
    },
    {
        "first": 'Thomas',
        "middle": '',
        "last": 'Truskett',
        "contributor": 'Truskett, T.',
        "scopus_ids": [6701644278],
        "clusters": ['Advanced Materials and Drug Delivery', 'Nanotechnology'],
    },
    {
        "first": 'Anish',
        "middle": '',
        "last": 'Tuteja',
        "contributor": 'Tuteja, A.',
        "scopus_ids": [6603587887],
        "clusters": ['Advanced Materials and Drug Delivery'],
    },
    {
        "first": 'J. Scott',
        "middle": 'S.',
        "last": 'VanEpps',
        "contributor": 'VanEpps, J. S.',
        "scopus_ids": [57211284617, 57222515420],
        "clusters": ['Nanotechnology'],
    },
    {
        "first": 'James',
        "middle": 'D.',
        "last": 'Weiland',
        "contributor": 'Weiland, J. D.',
        "scopus_ids": [7006322846],
        "clusters": ['Neural Engineering'],
    },
    {
        "first": 'Max',
        "middle": 'S.',
        "last": 'Wicha',
        "contributor": 'Wicha, M. S.',
        "scopus_ids": [7006067517, 57203044772, 57121076000],
        "clusters": ['Cell and Tissue Engineering', 'Single Cell Technologies'],
    },
    {
        "first": 'Matthew',
        "middle": '',
        "last": 'Willsey',
        "contributor": 'Willsey, M.',
        "scopus_ids": [55859799800],
        "clusters": ['Neural Engineering'],
    },
    {
        "first": 'Guizhi',
        "middle": '',
        "last": 'Zhu',
        "contributor": 'Zhu, G.',
        "scopus_ids": [8063118400, 57221993226],
        "clusters": ['BioInnovations in Brain Cancer'],
    },
]

# Flatten the readable roster into the structures used by the query and lookup.
# Every Scopus ID points back to the same faculty record, including secondary IDs.
FACULTY_BY_SCOPUS_ID = {
    scopus_id: faculty
    for faculty in BI_FACULTY_ROSTER
    for scopus_id in faculty["scopus_ids"]
}
bi_faculty = list(FACULTY_BY_SCOPUS_ID)
bi_faculty_set = set(bi_faculty)


# ---- Query Scopus directly, scoped to just this year -----------------------
print(f"Querying Scopus for BI-faculty publications in {TARGET_YEAR}...")
CHUNK = 15  # batch faculty IDs into a few OR-combined queries rather than one per person
docs_by_eid = {}
for i in range(0, len(bi_faculty), CHUNK):
    chunk_ids = bi_faculty[i:i + CHUNK]
    au_clause = " OR ".join(f"AU-ID({fid})" for fid in chunk_ids)
    query = f"({au_clause}) AND PUBYEAR IS {TARGET_YEAR}"
    try:
        s = ScopusSearch(query, view="COMPLETE", refresh=30)
        results = s.results or []
        print(f"  Chunk starting at {chunk_ids[0]}: {len(results)} result(s)")
        for doc in results:
            if doc.eid:
                docs_by_eid[doc.eid] = doc
    except Exception as e:
        print(f"  ScopusSearch failed for chunk starting at {chunk_ids[0]}: {e}")

print(f"Found {len(docs_by_eid)} unique BI-faculty publication(s) for {TARGET_YEAR}.")

# ---- Build records -----------------------------------------------------
records = []
skipped_no_doi = 0
skipped_no_date = 0

for eid, doc in docs_by_eid.items():
    if not doc.doi:
        skipped_no_doi += 1
        continue
    if not doc.coverDate:
        skipped_no_date += 1
        continue

    pub_date = date.fromisoformat(doc.coverDate)
    last_day = calendar.monthrange(pub_date.year, pub_date.month)[1]
    publish_date = date(pub_date.year, pub_date.month, last_day)
    link = f"https://doi.org/{doc.doi}"

    names = (doc.author_names or "").split(";")
    ids = (doc.author_ids or "").split(";")
    contributor_strs = []
    clusters = []
    for name_str, id_str in zip(names, ids):
        id_str = (id_str or "").strip()
        if not id_str.isdigit():
            continue
        faculty = FACULTY_BY_SCOPUS_ID.get(int(id_str))
        if faculty is None:
            continue

        # Use the canonical roster name so alternate Scopus profiles and byline
        # variations still produce one consistent contributor name.
        contributor_strs.append(faculty["contributor"])
        for cl in faculty["clusters"]:
            if cl not in clusters:
                clusters.append(cl)

    if not contributor_strs:
        # Shouldn't normally happen (the search itself was scoped to bi_faculty IDs),
        # but guards against an auid-format mismatch edge case.
        continue

    records.append({
        "Title": doc.title,
        "Date": pub_date.strftime("%B %Y"),
        "Journal": doc.publicationName,
        "Contributors": ", ".join(contributor_strs),
        "Link": link,
        "Clusters": ", ".join(clusters),
        "Facilities": "",
        "_PublishDateSort": publish_date,
        "Publish Date": publish_date.strftime("%B %d, %Y"),
    })

if skipped_no_doi:
    print(f"Skipped {skipped_no_doi} publication(s) with no DOI.")
if skipped_no_date:
    print(f"Skipped {skipped_no_date} publication(s) with no coverDate.")

if not records:
    raise SystemExit(f"No usable BI-faculty publications found for {TARGET_YEAR}.")

# ---- Assemble, sort, and write ---------------------------------------------
df = pd.DataFrame(records).sort_values("_PublishDateSort").reset_index(drop=True)

main_cols = ["Title", "Date", "Journal", "Contributors", "Link", "Clusters", "Facilities", "Publish Date"]
main_df = df[main_cols]
main_df.to_excel(OUT_MAIN, index=False)
print(f"\nWrote {OUT_MAIN} ({len(main_df)} rows).")

template_df = df[["Title", "Date", "Journal", "Contributors", "Link", "Clusters", "Publish Date"]].rename(
    columns={"Link": "Link to Publication"}
)
template_df.to_excel(OUT_TEMPLATE, index=False, engine="openpyxl")
print(f"Wrote {OUT_TEMPLATE} ({len(template_df)} rows).")
