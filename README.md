# GPO Migration: ENGR → CAD

Moves selected Group Policy Objects from the source domain (`ad.engr.wisc.edu`) into the
target domain (`ad.wisc.edu`). Each GPO is renamed with an `ENGR-` prefix and linked to the
matching OU under `OU=ENGR,OU=orgUnits,DC=ad,DC=wisc,DC=edu`.

It's built to be re-run safely. Run it once for what's ready today, then again as more
target OUs are built. Each run only adds what's new and never changes what's already there.

---

## Contents

| File | What it is |
|---|---|
| `GPO-Migrate.ps1` | Main script. `-Mode Export` runs on the source side, `-Mode Import` on the target side. |
| `build_ou_map.py` | Builds the source OU → target OU map from two OU exports (full-path matching). |
| `OU-Map.csv` | The current OU map, the output of `build_ou_map.py`. |
| `README.md` | This file. |

---

## Requirements

- **PowerShell 5.1+** with RSAT **Group Policy Management** and the **Active Directory** module, on both sides.
- **Export account:** read access to the source GPOs, OUs, and site configuration.
- **Import account:** rights to create GPOs in the target domain and to link GPOs on the target OUs.
- **Python 3.8+** (standard library only), only if you're rebuilding the OU map.
- A working folder on each side, `C:\Temp\GPO_Migration` by default (change it with `-WorkRoot`).

---

## How it works

```
 SOURCE (ad.engr.wisc.edu)                 TARGET (ad.wisc.edu)
 ─────────────────────────                 ─────────────────────
 GPO-InScope.csv ─┐
                  ├─ Export ─► Backups\             ┌─► Import ─► ENGR-<name> GPOs
 AD (read only) ──┘            manifest.csv   ─copy─┤             + links on mapped OUs
                               links.csv            │
                               MigrationTable-      ├── OU-Map.csv
                                 template.migtable  └── MigrationTable.migtable (edited)
                                                          │
                                                          ▼
                                                    import_results.csv
                                                    link_results.csv
```

**What gets migrated:**
- GPO settings, imported under the new name (UNC paths and principals are translated by the migration table)
- OU and domain-root links, including link order, enabled state, and enforced state
- GPO status (e.g. user settings disabled)

**What does not carry over. These are reported, and you handle them manually:**
- Site links
- WMI filters. The filter isn't created in the target, and the GPO's links are held disabled.
- Security filtering that narrows who the GPO applies to. Links are held disabled.
- GPO delegation (who can edit the GPO in the target). `Import-GPO` doesn't copy it.
- Block Inheritance on source OUs. It's noted in `link_results.csv` for you to check.

---

## Step by step

### 1. Prepare the inputs

**In-scope list:** In the inventory workbook, open the **Group Policy Inventory** sheet.
- Set `InScope` = `Yes` for each GPO to migrate. Leading and trailing spaces are ignored.
- Optionally, set `TargetGPOName` to use a name other than `ENGR-<name>`.
- Save the sheet as CSV UTF-8 to `C:\Temp\GPO_Migration\GPO-InScope.csv` on the source side.

Columns the script reads: `GPOName`, `InScope`, `TargetGPOName`. Any other columns are ignored.

**OU map:** see [Building the OU map](#building-the-ou-map). Only the import needs it.

### 2. Export (source side)

```powershell
.\GPO-Migrate.ps1 -Mode Export -DryRun     # lists what would be backed up; changes nothing
.\GPO-Migrate.ps1 -Mode Export
```

Export is read-only against AD. It writes these files to the working folder:

| File | Contents |
|---|---|
| `Backups\` | One GPO backup per in-scope GPO |
| `Reports\*.html` | Settings report per GPO, for review |
| `manifest.csv` | Source name → target name, backup ID, GPO status, WMI filter, security filtering |
| `links.csv` | One row per link: GPO, source OU, link order, enabled, enforced, block inheritance, link type |
| `not-found.csv` | In-scope names that don't exist in the source domain (renamed or deleted since the inventory) |
| `MigrationTable-template.migtable` | Every principal and UNC path the GPOs reference |
| `migration.log` | Full log. Appended on every run, both modes. |

Review `manifest.csv` and `not-found.csv` before going on.

### 3. Build the migration table

1. Open `MigrationTable-template.migtable` in the **GPMC Migration Table Editor** (in GPMC, right-click *Group Policy Objects* → *Open Migration Table Editor*).
2. For each source principal (`...@ad.engr.wisc.edu`) and UNC path, set the target-side equivalent.
3. Save it as **`MigrationTable.migtable`** in the same folder.

If `MigrationTable.migtable` isn't there, the import still runs, but ENGR principals and paths come across unchanged and won't resolve in CAD. The log warns about this.

Built-in names (`Administrators`, `Authenticated Users`, `LOCAL SERVICE`, and so on) usually stay *Same as source*. Bare SIDs (`S-1-5-21-...`) are deleted accounts. Map them to nothing, or to a replacement group.

### 4. Copy to the target side

Copy the entire `C:\Temp\GPO_Migration` folder to the target machine, and put `OU-Map.csv` in it.

### 5. Import dry run (target side)

```powershell
.\GPO-Migrate.ps1 -Mode Import -DryRun -OnlyLinked
```

Nothing changes in AD. Review:
- `import_results.csv`: what happens to each GPO
- `link_results.csv`: every link, where it would land, and why any are skipped

Fix the OU map or the in-scope list, and dry-run again until the results are what you expect.

### 6. Import

```powershell
.\GPO-Migrate.ps1 -Mode Import -OnlyLinked
```

### 7. After the import

- In `link_results.csv`, filter `LinkEnabled` = `No`. Those GPOs need their WMI filter recreated or their security filtering set in the target before you enable the link.
- Check the delegation on the new GPOs in GPMC.
- Spot-check a few GPOs' settings against the HTML reports in `Reports\`.

### Re-running as more OUs are built

1. Rebuild `OU-Map.csv` (next section).
2. Run the import again, dry run first.

GPOs already created are left as they are, and links that already exist aren't touched. Only newly mappable links, and with `-OnlyLinked` newly linkable GPOs, are added. Re-run the export only if the in-scope list changes or the source GPOs change.

---

## Building the OU map

The map must be built by **full path**, not by OU name. Many OU names repeat (`PI`, `Grad`, `Computers`, `Lab`...). A name-based lookup silently sends GPOs to the wrong OU.

The script refuses any source OU that appears on more than one row of the map (`SKIPPED-AmbiguousMap`), but it can't catch every name-based mistake. Use `build_ou_map.py`.

**1. Export both OU trees**, one on each side:

```powershell
# Source side
Get-ADOrganizationalUnit -Filter * -Server ad.engr.wisc.edu -Properties canonicalName |
  Select-Object name, distinguishedName, canonicalName |
  Export-Csv ENGR-OU.Structure.csv -NoTypeInformation

# Target side
Get-ADOrganizationalUnit -Filter * -Server ad.wisc.edu `
  -SearchBase 'OU=ENGR,OU=orgUnits,DC=ad,DC=wisc,DC=edu' -Properties canonicalName |
  Select-Object name, distinguishedName, canonicalName |
  Export-Csv CAD-ENGR.OUStructure.csv -NoTypeInformation
```

**2. Build the map:**

```
python build_ou_map.py ENGR-OU.Structure.csv CAD-ENGR.OUStructure.csv OU-Map.csv GPO-InScope.csv
```

The last argument is optional. It adds `GPO Links` and `In-Scope GPO Links` counts, so you can see which OUs matter.

**3. Review `OU-Map.csv`.** Each row has a `Match` value:

| Match | Meaning | `CAD OU` filled? |
|---|---|---|
| `Exact` | Same path under the renamed department (`Mechanical Engineering/1st Floor` → `ME/1st Floor`) | Yes |
| `Inserted level` | Same path with 1–2 extra target levels (`CAE/Computers/ECB` → `CAE/Computers/CAE Labs/ECB`). Only used when exactly one target fits. **Review these.** | Yes |
| `Ambiguous` | More than one target fits, or the target is already taken; candidates are listed | No |
| `None` | No matching target OU exists yet | No |

The GPO script only uses rows with a `CAD OU`. To map an OU by hand, paste the target DN into `CAD OU`. To stop an OU being linked, clear its `CAD OU`.

**Department translation** is the `DEPT_MAP` block at the top of `build_ou_map.py` (e.g. `'mechanical engineering': ['ME']`). When a department appears under a new name in CAD, add or adjust it there. When a department has two candidate trees (`CBE` and `CBENG` today), the first one listed wins, and rows that exist in both are noted.

---

## Reference

### Parameters (`GPO-Migrate.ps1`)

| Parameter | Default | Purpose |
|---|---|---|
| `-Mode` | *(required)* | `Export` (source side) or `Import` (target side) |
| `-DryRun` | off | Do every lookup and write the result files, but change nothing |
| `-WorkRoot` | `C:\Temp\GPO_Migration` | Working folder for inputs and outputs |
| `-InScopeCsv` | `<WorkRoot>\GPO-InScope.csv` | In-scope list |
| `-OUMapCsv` | `<WorkRoot>\OU-Map.csv` | OU map |
| `-SourceDomain` | `ad.engr.wisc.edu` | Source domain |
| `-TargetDomain` | `ad.wisc.edu` | Target domain |
| `-NamePrefix` | `ENGR-` | Prefix for target GPO names (not doubled if the name already has it) |
| `-OnlyLinked` | off | Import only GPOs that have at least one mapped link. **Recommended**: it keeps CAD from filling up with unlinked GPOs. |
| `-LinkState` | `AsSource` | `AsSource` keeps each link's enabled state. `Disabled` creates every link disabled, for enabling per wave. |
| `-OverwriteExisting` | off | Re-import settings into a GPO this process created on an earlier run. Never applies to GPOs it didn't create. |

The column names the script reads are set in the `$Columns` block near the top of the script.

### Import outcomes (`import_results.csv`, `Action` column)

| Action | Meaning |
|---|---|
| `Created` / `WOULD-Create` | New GPO created (or would be, in a dry run) |
| `Exists-NotChanged` | Created by an earlier run; settings left as-is. Missing links are still added. |
| `Overwritten` / `WOULD-Overwrite` | Settings re-imported (`-OverwriteExisting`) |
| `SKIPPED-NameInUse` | A GPO with the target name already exists and wasn't created by this process. It's never touched, and its links are skipped. |
| `SKIPPED-NoMappedLinks` | `-OnlyLinked` is set and none of the GPO's links map to a target OU yet |
| `FAILED` | See the `Error` column and `migration.log` |

GPOs this process creates are tagged in their description with `[ENGR-MIG {source GPO GUID}]`. That tag is how later runs recognize them. **Don't remove it.**

### Link outcomes (`link_results.csv`, `Action` column)

| Action | Meaning |
|---|---|
| `Linked` / `WOULD-Link` | Link created (or would be) on `TargetOU` |
| `AlreadyLinked-NotChanged` | The link already exists; its state isn't changed |
| `SKIPPED-Unmapped` | The source OU has no `CAD OU` in the map |
| `SKIPPED-AmbiguousMap` | The source OU appears on more than one map row; the map was probably built by name |
| `SKIPPED-TargetOUNotFound` | The mapped target OU doesn't exist in the target domain |
| `SKIPPED-Duplicate` | Two source OUs map to the same target OU for the same GPO; only one link is made |
| `SKIPPED-GPONotImported` | The GPO wasn't imported this run (see `import_results.csv`) |
| `SKIPPED-SiteLink` | Site links aren't migrated |
| `FAILED` | See the `Note` column |

The `LinkEnabled` column shows the state each link was or will be created with, and `Note` explains any link held disabled.

### When links are held disabled

A link is created disabled, whatever its source state, when either of these is true:
- The GPO has a **WMI filter**.
- Its **security filtering narrows who it applies to**: Authenticated Users isn't allowed to apply it, or there's a deny entry.

Extra groups alongside Authenticated Users don't count, because the GPO still applies to everyone. Recreate the filter in the target, then enable the link.

### Link order

On each target OU, migrated links keep their relative order from the source. They're added after any links already on that OU, so existing target GPOs keep precedence. Review the order in GPMC if it matters for that OU.

### What stops a run

The whole run only stops for setup problems:
- missing RSAT modules
- missing input files (in-scope list, manifest, links, or OU map)
- a required column missing from an input

Everything else is handled one GPO or one link at a time and recorded in the result files.

---

## Troubleshooting

| Symptom | Likely cause / fix |
|---|---|
| Many `SKIPPED-Unmapped` | The target OUs aren't built yet, or the map needs rebuilding. Re-export the target OU tree and re-run `build_ou_map.py`. |
| `SKIPPED-AmbiguousMap` | The OU map was built by OU name. Rebuild it with `build_ou_map.py`. |
| `SKIPPED-NameInUse` | Someone already created a GPO with that name. Set a different `TargetGPOName` and re-export, or rename or remove the existing GPO. |
| GPO names in `not-found.csv` | Renamed or deleted in the source since the inventory. Update the inventory. |
| Principals like `ENGR\group` in imported GPO settings | `MigrationTable.migtable` was missing or incomplete. Fix it, then re-run the import with `-OverwriteExisting`. |
| "Could not build migration table template" | The GPMC COM object isn't available on that machine. Build the table in GPMC's Migration Table Editor from the backups instead. |

---

## Reusing for another department or domain

- **Script:** set `-SourceDomain`, `-TargetDomain`, and `-NamePrefix`, and adjust `$Columns` if your spreadsheet headers differ.
- **Description tag:** change `$MarkerPrefix` so runs for different departments don't recognize each other's GPOs.
- **OU map:** change `DEPT_MAP` and `ROOT_MAP` in `build_ou_map.py`. `tgt_path()` assumes the target tree sits at `<domain>/orgUnits/<dept root>/...`, so adjust it if yours is laid out differently.
