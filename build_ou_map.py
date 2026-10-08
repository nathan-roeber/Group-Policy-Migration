"""
Build a full-path OU map: ENGR source OU -> CAD target OU (ad.wisc.edu/orgUnits/ENGR/...).

Matching, per source OU:
  Exact           same path under the translated department (Mechanical Engineering -> ME, ...)
  Inserted level  same segments in the same order and same leaf, with 1-2 extra levels on
                  the target side (CAE/Computers/ECB -> CAE/Computers/CAE Labs/ECB); only
                  used when exactly one target fits
  Ambiguous       more than one target fits; candidates listed, CAD OU left blank
  None            no target yet (e.g. Engineering Physics, CEE, ECE aren't built in CAD)

Only Exact and Inserted-level rows get a CAD OU, so anything else stays unmapped for the
GPO script. Column names match what GPO-Migrate.ps1 reads ("ENGR OU", "CAD OU").
"""
import csv
import collections
import re
import sys

ENGR_CSV, CAD_CSV, OUT_CSV = sys.argv[1], sys.argv[2], sys.argv[3]
GPO_INVENTORY_CSV = sys.argv[4] if len(sys.argv) > 4 else None   # optional: adds link counts

# Source department (first level under orgUnits, or a domain-root OU) -> target department(s)
# under OU=ENGR, in order of preference. First listed wins when both have an exact match.
DEPT_MAP = {
    'mechanical engineering': ['ME'],
    'cbe': ['CBE', 'CBENG'],
    'cae': ['CAE'],
    'bme': ['BME'],
    'engine resource center': ['ERC'],
    'managed desktops': ['Managed Desktop'],
    'mechlab': ['MechLab'],
    'hsx': ['HSX'],
    'tops': ['TOPS'],
    'engineering media services': ['EMS'],
    'industrial engineering': ['IE', 'INPRO'],
    'engineering physics': ['EP'],
    'ceete': ['CEETE'],
    'ece': ['ECE'],
    'ece admin': ['ECE Admin'],
    'makerspace': ['Makerspace'],
    'msae': ['MSAE'],
}
# OUs at the domain root (outside orgUnits) -> target department. Anything not listed is unmapped.
ROOT_MAP = {
    'newcomputers': ['NewComputers'],
}
MAX_INSERTED = 2

def src_path(canonical):
    """(path from the department down, True if under orgUnits)"""
    p = canonical.split('/')[1:]                  # drop the domain
    return (p[1:], True) if p and p[0] == 'orgUnits' else (p, False)

def tgt_path(canonical):
    return canonical.split('/')[3:]               # drop ad.wisc.edu/orgUnits/ENGR

def is_subsequence(short, long):
    it = iter(long)
    return all(any(x == y for y in it) for x in short)

S = list(csv.DictReader(open(ENGR_CSV, encoding='utf-8-sig')))
T = list(csv.DictReader(open(CAD_CSV, encoding='utf-8-sig')))
by_path = {'/'.join(tgt_path(t['canonicalName'])).lower(): t for t in T}
by_dept = collections.defaultdict(list)
for t in T:
    tp = tgt_path(t['canonicalName'])
    if tp:
        by_dept[tp[0].lower()].append(tp)

# Optional: how many GPO links (all GPOs / in-scope GPOs) point at each source OU
links_all, links_in = collections.Counter(), collections.Counter()
if GPO_INVENTORY_CSV:
    for r in csv.DictReader(open(GPO_INVENTORY_CSV, encoding='utf-8-sig')):
        inscope = (r.get('InScope') or '').strip().lower() == 'yes'
        for cn in (r.get('LinkedOUs') or '').split('|'):
            cn = cn.strip().lower()
            if cn:
                links_all[cn] += 1
                if inscope:
                    links_in[cn] += 1

out = []
claimed_exact = set()
pending_inserted = []   # (row, candidate) resolved after all exact matches are known
for s in S:
    sp, under_org = src_path(s['canonicalName'])
    row = {'Source OU Name': s['name'], 'ENGR OU': s['distinguishedName'], 'CAD OU': '',
           'Match': 'None', 'Source Path': '/'.join(sp), 'Target Path': '', 'Candidates': '', 'Note': '',
           'GPO Links': links_all[s['canonicalName'].lower()],
           'In-Scope GPO Links': links_in[s['canonicalName'].lower()]}
    depts = ((DEPT_MAP if under_org else ROOT_MAP).get(sp[0].lower(), [])) if sp else []
    if not depts:
        row['Note'] = f"no target department for '{sp[0] if sp else ''}'"
        out.append(row); continue
    rest = [x.lower() for x in sp[1:]]

    # Exact
    exact = [d for d in depts if '/'.join([d] + sp[1:]).lower() in by_path]
    if exact:
        t = by_path['/'.join([exact[0]] + sp[1:]).lower()]
        row.update({'CAD OU': t['distinguishedName'], 'Match': 'Exact', 'Target Path': '/'.join(tgt_path(t['canonicalName']))})
        claimed_exact.add(t['distinguishedName'].lower())
        if len(exact) > 1:
            row['Note'] = f"also exists under {', '.join(exact[1:])}"
        out.append(row); continue

    # Inserted level(s): same leaf, source segments in order, 1-2 extra target levels
    cands = []
    if rest:
        for d in depts:
            for tp in by_dept.get(d.lower(), []):
                tr = [x.lower() for x in tp[1:]]
                if tr and tr[-1] == rest[-1] and 0 < len(tr) - len(rest) <= MAX_INSERTED and is_subsequence(rest, tr):
                    cands.append('/'.join(tp))
    if len(cands) == 1:
        pending_inserted.append((row, cands[0]))
    elif cands:
        row.update({'Match': 'Ambiguous', 'Candidates': ' | '.join(sorted(cands))})
    else:
        row['Note'] = f"nothing under ENGR/{'|'.join(depts)} matches"
    out.append(row)

# Inserted-level matches can't take a target an exact match owns, or share one with each other
claims = collections.Counter(c.lower() for _, c in pending_inserted)
for row, cand in pending_inserted:
    t = by_path[cand.lower()]
    if t['distinguishedName'].lower() in claimed_exact:
        row.update({'Match': 'Ambiguous', 'Candidates': cand, 'Note': 'target already matched exactly by another source OU'})
    elif claims[cand.lower()] > 1:
        row.update({'Match': 'Ambiguous', 'Candidates': cand, 'Note': 'more than one source OU fits this target'})
    else:
        row.update({'CAD OU': t['distinguishedName'], 'Match': 'Inserted level', 'Target Path': cand})

with open(OUT_CSV, 'w', newline='', encoding='utf-8-sig') as f:
    w = csv.DictWriter(f, fieldnames=list(out[0].keys()))
    w.writeheader()
    w.writerows(out)

c = collections.Counter(r['Match'] for r in out)
print(f"{len(out)} source OUs -> {OUT_CSV}")
for k in ('Exact', 'Inserted level', 'Ambiguous', 'None'):
    print(f"  {k:<15} {c[k]}")
used = collections.Counter(r['CAD OU'].lower() for r in out if r['CAD OU'])
dup = {k: v for k, v in used.items() if v > 1}
print(f"  target OUs used by more than one source OU: {len(dup)}")
