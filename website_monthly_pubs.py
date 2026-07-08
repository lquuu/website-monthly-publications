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
# - Surnames are normalized (case, spaces, hyphens stripped) before matching against
#   the cluster roster below, so "VanEpps"/"Vanepps"/"Van Epps" etc. all match.
# - "Nagrath" and "Schwendeman" are disambiguated by first initial (two people each).
# - Any surname encountered that ISN'T in the roster gets flagged in the printed
#   output at the end, rather than silently left blank -- check that list and let
#   me know the missing clusters.
# - Publications with no DOI are skipped (can't build a Link), same as the original
#   collection script's behavior.

import re
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

bi_faculty = [
    55016188700, 24068798200, 12779584500, 57195268661, 26021803700,
    7004361447, 15033338100, 55465936200, 8072516000, 57201972955,
    7202812560, 9536007500, 57203325389, 22135180200, 6603555265,
    7402161971, 16042824100, 57200610670, 36930755200, 57189033142,
    15035854800, 7006008866, 55975937900, 35726004100, 8577389100,
    23101128900, 57189212345, 55885015000, 35243347700, 6506939515,
    57222518153, 7003309460, 7005328686, 35307875200, 7005915503,
    7101689432, 6701644278, 6603587887, 57211284617, 57222515420,
    7006322846, 7006067517, 57203044772, 55859799800,
]
bi_faculty_set = set(bi_faculty)

# ---- Cluster roster -------------------------------------------------------
# Keys are normalized surnames (lowercase, spaces/hyphens stripped).
CLUSTER_ROSTER = {
    "aguilar": ["Cell and Tissue Engineering", "Single Cell Technologies"],
    "bruns": ["Neural Engineering"],
    "chestek": ["Neural Engineering"],
    "chun": ["Cell and Tissue Engineering"],
    "glotzer": ["Nanotechnology"],
    "greineder": ["Nanotechnology"],
    "hara": ["BioInnovations in Brain Cancer"],
    "hu": ["Cell and Tissue Engineering"],
    "keller": ["Cell and Tissue Engineering", "Single Cell Technologies"],
    "kim": ["Advanced Materials and Drug Delivery"],
    "kotov": ["Nanotechnology", "Single Cell Technologies"],
    "lahann": ["Advanced Materials and Drug Delivery", "BioInnovations in Brain Cancer", "Single Cell Technologies"],
    "larson": ["Nanotechnology"],
    "lempka": ["Neural Engineering"],
    "lesherperez": ["Cell and Tissue Engineering"],
    "liu": ["Advanced Materials and Drug Delivery", "Nanotechnology", "Cell and Tissue Engineering"],
    "lombaert": ["Cell and Tissue Engineering"],
    "luker": ["Cell and Tissue Engineering", "Single Cell Technologies"],
    "min": ["Cell and Tissue Engineering", "BioInnovations in Brain Cancer", "Advanced Materials and Drug Delivery", "Nanotechnology"],
    "moon": ["BioInnovations in Brain Cancer", "Advanced Materials and Drug Delivery"],
    "opri": ["Neural Engineering"],
    "penafrancesch": ["Nanotechnology"],
    "rosenzweig": ["Cell and Tissue Engineering", "Single Cell Technologies"],
    "shea": ["Cell and Tissue Engineering"],
    "shin": ["Advanced Materials and Drug Delivery", "Cell and Tissue Engineering", "Single Cell Technologies", "Nanotechnology"],
    "solomon": ["Nanotechnology"],
    "stacey": ["Neural Engineering"],
    "tessier": ["BioInnovations in Brain Cancer", "Advanced Materials and Drug Delivery", "Nanotechnology"],
    "truskett": ["Advanced Materials and Drug Delivery", "Nanotechnology"],
    "tuteja": ["Advanced Materials and Drug Delivery"],
    "vanepps": ["Nanotechnology"],
    "weiland": ["Neural Engineering"],
    "wicha": ["Cell and Tissue Engineering", "Single Cell Technologies"],
    "willsey": ["Neural Engineering"],
    "eniolaadefeso": [],  # TODO: Luciana to provide cluster(s)
}

# Ambiguous surnames -- two BI faculty share these, disambiguated by first initial.
CLUSTER_ROSTER_BY_INITIAL = {
    ("nagrath", "d"): ["BioInnovations in Brain Cancer", "Cell and Tissue Engineering", "Single Cell Technologies"],
    ("nagrath", "s"): ["Cell and Tissue Engineering", "Single Cell Technologies"],
    ("schwendeman", "a"): ["Nanotechnology", "BioInnovations in Brain Cancer"],
    ("schwendeman", "s"): ["BioInnovations in Brain Cancer", "Advanced Materials and Drug Delivery"],
}
AMBIGUOUS_SURNAMES = {"nagrath", "schwendeman"}


def normalize_surname(surname):
    return re.sub(r"[^a-z]", "", (surname or "").lower())


def get_clusters(surname, initials_or_given_name):
    s = normalize_surname(surname)
    if s in AMBIGUOUS_SURNAMES:
        initial = (initials_or_given_name or "").strip()[:1].lower()
        return CLUSTER_ROSTER_BY_INITIAL.get((s, initial))
    return CLUSTER_ROSTER.get(s)


def parse_author_name(name_str):
    """ScopusSearch's author_names entries come as 'Surname, Initials', e.g.
    'Kotov, N.A.'. Split into (surname, initials)."""
    if not name_str:
        return "", ""
    if "," in name_str:
        surname, initials = name_str.split(",", 1)
    else:
        surname, initials = name_str, ""
    return surname.strip(), initials.strip()


def format_contributor(surname, given_or_initials):
    """Handles both already-abbreviated initials ('N.A.') and full given names
    ('PETER M.', 'Omolola') -- takes the first letter of each letter-run either way."""
    given_or_initials = (given_or_initials or "").strip()
    if not given_or_initials:
        return surname
    tokens = re.findall(r"[A-Za-z]+", given_or_initials)
    if not tokens:
        return surname
    formatted_initials = ". ".join(t[0].upper() for t in tokens) + "."
    return f"{surname}, {formatted_initials}"


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
missing_surnames = set()
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
        if not id_str.isdigit() or int(id_str) not in bi_faculty_set:
            continue
        surname, initials = parse_author_name(name_str)
        contributor_strs.append(format_contributor(surname, initials))
        c = get_clusters(surname, initials)
        if c is None:
            missing_surnames.add(surname)
        elif c:
            for cl in c:
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

if missing_surnames:
    print("\nNOTE: these surnames weren't found in the cluster roster -- "
          "Clusters left blank for them. Let me know their clusters:")
    for s in sorted(missing_surnames):
        print(f"  - {s}")

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