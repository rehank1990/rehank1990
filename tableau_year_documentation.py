#!/usr/bin/env python3
"""
Tableau 2026 reports -- download + documentation pack
=====================================================

Finds every workbook published to Tableau Server in a given year, downloads it,
reads what is inside, and writes a multi-sheet Excel pack for consultants.

Output:

    C:\\tb\\2026\\
        Finance\\
            Loan Portfolio Summary.twb
        Operations\\
            Branch Traffic.twb
        Tableau_2026_Report_Documentation.xlsx
        _issues.csv

The Excel pack has one sheet per kind of detail:

    Summary              counts and totals for the whole set
    Workbooks            one row per report: project, owner, dates, object counts
    Connections          every database connection: type, server, database, schema
    Tables & Queries     physical tables, custom SQL, and stored procedures
    Calculated Fields    every calculation with its formula, LOD/table-calc flags
    Parameters           parameters with data type and current value
    Views                worksheets, dashboards, and stories per workbook
    Published Data Srcs  which workbooks depend on which published data sources
    Issues               anything that could not be downloaded or parsed

WHAT "PUBLISHED IN 2026" MEANS
------------------------------
Tableau tracks two dates. By default this uses the created date -- when the
workbook first appeared on the server. Republishing an older workbook updates
its modified date but not its created date, so:

    date_field="created"   first published in 2026            (default)
    date_field="modified"  changed in 2026, whenever created
    date_field="either"    created OR modified in 2026        (widest net)

RUNNING IT
----------
Notebook:

    import os, getpass
    os.environ["TABLEAU_USERNAME"] = "rhkhan"
    os.environ["TABLEAU_PASSWORD"] = getpass.getpass("Tableau password: ")

    from tableau_year_documentation import document_year
    document_year(year=2026, out=r"C:\\tb\\2026", dry_run=True)

Terminal:

    python tableau_year_documentation.py --year 2026 --out C:\\tb\\2026 --dry-run

Extracts are EXCLUDED by default: the documentation comes from the workbook XML,
and skipping extracts makes this run in minutes instead of hours. Pass
include_extracts=True if the consultants need the data too.

Configuration via environment variables:

    TABLEAU_SERVER      dc-p-tableau     (scheme optional -- added automatically)
    TABLEAU_SITE        "" for the Default site
    TABLEAU_TOKEN_NAME  + TABLEAU_TOKEN_SECRET   (preferred)
    TABLEAU_USERNAME    + TABLEAU_PASSWORD       (fallback)
    TABLEAU_CA_BUNDLE   path to internal CA, or "false" to skip verification

Requires: pip install tableauserverclient pandas openpyxl
"""

from __future__ import annotations

import csv
import os
import re
import sys
import time
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

import pandas as pd
import tableauserverclient as TSC

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

SERVER_HOST = os.environ.get("TABLEAU_SERVER", "dc-p-tableau")
SITE_ID = os.environ.get("TABLEAU_SITE", "")
CA_BUNDLE = os.environ.get("TABLEAU_CA_BUNDLE", "")

PAGE_SIZE = 1000
DOWNLOAD_TIMEOUT_SECONDS = 1800

# Excel rejects any cell over 32,767 characters, and long custom SQL will hit
# that. Truncate with a marker so it's obvious the text was cut.
EXCEL_CELL_LIMIT = 32_000

LOD_PATTERN = re.compile(r"\{\s*(FIXED|INCLUDE|EXCLUDE)", re.IGNORECASE)
TABLE_CALC_PATTERN = re.compile(
    r"\b(WINDOW_\w+|RUNNING_\w+|LOOKUP|INDEX|RANK\w*|TOTAL|FIRST|LAST|PREVIOUS_VALUE)\s*\(",
    re.IGNORECASE,
)

ILLEGAL_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
RESERVED_NAMES = {
    "CON", "PRN", "AUX", "NUL",
    *(f"COM{i}" for i in range(1, 10)),
    *(f"LPT{i}" for i in range(1, 10)),
}
MAX_NAME_LEN = 120


# ---------------------------------------------------------------------------
# Connection
# ---------------------------------------------------------------------------


def resolve_server_url(host: str) -> str:
    """Turn a bare hostname into a URL the REST API library will accept."""
    host = host.strip().rstrip("/")
    if host.startswith(("http://", "https://")):
        return host

    import socket
    import ssl

    hostname = host.split("/")[0]

    try:
        context = ssl.create_default_context()
        context.check_hostname = False
        context.verify_mode = ssl.CERT_NONE
        with socket.create_connection((hostname, 443), timeout=5) as raw:
            with context.wrap_socket(raw, server_hostname=hostname):
                return f"https://{host}"
    except Exception:  # noqa: BLE001
        pass

    try:
        with socket.create_connection((hostname, 80), timeout=5):
            print(f"Note: {hostname} answered on HTTP, not HTTPS.")
            return f"http://{host}"
    except Exception:  # noqa: BLE001
        pass

    print(
        f"Warning: could not reach {hostname} on port 443 or 80. "
        "Defaulting to HTTPS -- check DNS/VPN if sign-in fails."
    )
    return f"https://{host}"


def connect() -> TSC.Server:
    """Sign in with a token if configured, otherwise username and password."""
    token_name = os.environ.get("TABLEAU_TOKEN_NAME")
    token_secret = os.environ.get("TABLEAU_TOKEN_SECRET")
    username = os.environ.get("TABLEAU_USERNAME")
    password = os.environ.get("TABLEAU_PASSWORD")

    if token_name and token_secret:
        auth = TSC.PersonalAccessTokenAuth(token_name, token_secret, site_id=SITE_ID)
        method = f"token '{token_name}'"
    elif username and password:
        auth = TSC.TableauAuth(username, password, site_id=SITE_ID)
        method = f"user '{username}'"
    else:
        raise RuntimeError(
            "No credentials found. Set either TABLEAU_TOKEN_NAME + "
            "TABLEAU_TOKEN_SECRET, or TABLEAU_USERNAME + TABLEAU_PASSWORD."
        )

    url = resolve_server_url(SERVER_HOST)
    server = TSC.Server(url, use_server_version=True)

    http_options: dict = {"timeout": DOWNLOAD_TIMEOUT_SECONDS}
    if CA_BUNDLE:
        http_options["verify"] = False if CA_BUNDLE.lower() == "false" else CA_BUNDLE
    server.add_http_options(http_options)

    server.auth.sign_in(auth)
    print(f"Connected to {url} as {method} (REST API {server.version})")
    return server


def fetch_all(endpoint_get) -> list:
    """Page through a TSC .get() endpoint and return every item."""
    items, page = [], 1
    while True:
        opts = TSC.RequestOptions(pagenumber=page, pagesize=PAGE_SIZE)
        batch, pagination = endpoint_get(opts)
        items.extend(batch)
        if not batch or len(items) >= pagination.total_available:
            return items
        page += 1


# ---------------------------------------------------------------------------
# Naming and small helpers
# ---------------------------------------------------------------------------


def safe_name(name: str) -> str:
    """Turn a Tableau content name into something a filesystem will accept."""
    cleaned = ILLEGAL_CHARS.sub("_", name or "unnamed").strip().rstrip(". ")
    if cleaned.upper().split(".")[0] in RESERVED_NAMES:
        cleaned += "_"
    if len(cleaned) > MAX_NAME_LEN:
        cleaned = cleaned[:MAX_NAME_LEN].rstrip()
    return cleaned or "unnamed"


def build_project_paths(projects) -> dict:
    """Map project id -> list of folder names, root first."""
    by_id = {p.id: p for p in projects}
    paths = {}
    for proj in projects:
        parts, node, guard = [], proj, 0
        while node is not None and guard < 20:
            parts.append(safe_name(node.name))
            node = by_id.get(node.parent_id) if node.parent_id else None
            guard += 1
        paths[proj.id] = list(reversed(parts))
    return paths


def unique_path(directory: Path, stem: str, suffix: str) -> Path:
    candidate = directory / f"{stem}{suffix}"
    if not candidate.exists():
        return candidate
    for counter in range(2, 1000):
        candidate = directory / f"{stem} ({counter}){suffix}"
        if not candidate.exists():
            return candidate
    raise RuntimeError(f"Could not find a free filename for {stem}{suffix}")


def mb(size_bytes) -> float | None:
    return round(size_bytes / 1_048_576, 2) if size_bytes else None


def fmt_date(value) -> str:
    if not value:
        return ""
    return value if isinstance(value, str) else value.strftime("%Y-%m-%d")


def clean_field_name(raw: str) -> str:
    """Strip Tableau's [brackets] and internal prefixes from a field name."""
    text = (raw or "").strip()
    text = re.sub(r"^\[|\]$", "", text)
    return text


def clip(text: str, limit: int = EXCEL_CELL_LIMIT) -> str:
    """Keep a cell within Excel's character ceiling."""
    text = (text or "").strip()
    if len(text) <= limit:
        return text
    return text[:limit] + f"\n... [truncated, {len(text)} characters total]"


# ---------------------------------------------------------------------------
# Workbook XML analysis
# ---------------------------------------------------------------------------


def open_workbook_xml(package: Path) -> ET.Element:
    """Return the root element of a .twb, or of the .twb inside a .twbx."""
    if package.suffix.lower() == ".twbx":
        with zipfile.ZipFile(package) as archive:
            name = next(n for n in archive.namelist() if n.lower().endswith(".twb"))
            return ET.fromstring(archive.read(name))
    return ET.parse(package).getroot()


def analyze_workbook(root: ET.Element, wb_name: str) -> dict:
    """Pull every documentation detail out of one workbook's XML.

    Returns dicts of lists, one per Excel sheet, all tagged with the workbook
    name so the sheets can be filtered independently.
    """
    connections, relations, calcs, parameters, views, published = [], [], [], [], [], []

    for ds in root.iter("datasource"):
        ds_name = ds.get("caption") or ds.get("name") or ""

        # Parameters live in a pseudo-datasource named "Parameters".
        if ds.get("name") == "Parameters":
            for column in ds.iter("column"):
                parameters.append(
                    {
                        "Workbook": wb_name,
                        "Parameter": clean_field_name(column.get("caption") or column.get("name")),
                        "Data Type": column.get("datatype", ""),
                        "Current Value": column.get("value", ""),
                        "Allowable Values": column.get("param-domain-type", ""),
                    }
                )
            continue

        # A published data source is referenced by repository-location, whose id
        # is the data source's content URL on the server.
        location = ds.find("repository-location")
        if location is not None:
            published.append(
                {
                    "Workbook": wb_name,
                    "Published Data Source": ds_name,
                    "Content URL": location.get("id", ""),
                    "Site": location.get("site", ""),
                }
            )

        # Database connections. "federated" is Tableau's internal wrapper, not a
        # real database, so it's skipped.
        for conn in ds.iter("connection"):
            conn_class = conn.get("class", "")
            if conn_class in ("federated", ""):
                continue
            connections.append(
                {
                    "Workbook": wb_name,
                    "Data Source": ds_name,
                    "Connection Type": conn_class,
                    "Server": conn.get("server", ""),
                    "Port": conn.get("port", ""),
                    "Database": conn.get("dbname", ""),
                    "Schema": conn.get("schema", ""),
                    "Warehouse": conn.get("warehouse", ""),
                    "Role": conn.get("role", ""),
                    "Service": conn.get("service", ""),
                    "Filename": conn.get("filename", ""),
                    "Credentials Embedded": bool(conn.get("username")),
                }
            )

        # Tables, custom SQL, and stored procedures.
        for rel in ds.iter("relation"):
            rel_type = rel.get("type", "")
            if rel_type == "table":
                relations.append(
                    {
                        "Workbook": wb_name,
                        "Data Source": ds_name,
                        "Object Type": "Table",
                        "Name / Alias": rel.get("name", ""),
                        "Table or Query": clip(rel.get("table", "")),
                    }
                )
            elif rel_type == "text" and (rel.text or "").strip():
                relations.append(
                    {
                        "Workbook": wb_name,
                        "Data Source": ds_name,
                        "Object Type": "Custom SQL",
                        "Name / Alias": rel.get("name", ""),
                        "Table or Query": clip(rel.text),
                    }
                )
            elif rel_type == "stored-proc":
                relations.append(
                    {
                        "Workbook": wb_name,
                        "Data Source": ds_name,
                        "Object Type": "Stored Procedure",
                        "Name / Alias": rel.get("name", ""),
                        "Table or Query": clip(rel.get("stored-proc", "")),
                    }
                )

        # Calculated fields. The formula sits on a <calculation> inside a column.
        for column in ds.iter("column"):
            calc = column.find("calculation")
            if calc is None or calc.get("class") != "tableau":
                continue
            formula = calc.get("formula", "") or ""
            calcs.append(
                {
                    "Workbook": wb_name,
                    "Data Source": ds_name,
                    "Field": clean_field_name(column.get("caption") or column.get("name")),
                    "Data Type": column.get("datatype", ""),
                    "Role": column.get("role", ""),
                    "Is LOD": bool(LOD_PATTERN.search(formula)),
                    "Is Table Calc": bool(TABLE_CALC_PATTERN.search(formula)),
                    "Formula": clip(formula),
                }
            )

    # Worksheets, dashboards, stories.
    for sheet in root.iter("worksheet"):
        used = {d.get("name") for d in sheet.iter("datasource") if d.get("name")}
        used.discard("Parameters")
        views.append(
            {
                "Workbook": wb_name,
                "Object": sheet.get("name", ""),
                "Type": "Worksheet",
                "Data Sources Used": len(used),
                "Is Blend": len(used) > 1,
            }
        )
    for dash in root.iter("dashboard"):
        views.append(
            {
                "Workbook": wb_name,
                "Object": dash.get("name", ""),
                "Type": "Dashboard",
                "Data Sources Used": "",
                "Is Blend": "",
            }
        )
    for story in root.iter("story"):
        views.append(
            {
                "Workbook": wb_name,
                "Object": story.get("name", ""),
                "Type": "Story",
                "Data Sources Used": "",
                "Is Blend": "",
            }
        )

    return {
        "connections": connections,
        "relations": relations,
        "calcs": calcs,
        "parameters": parameters,
        "views": views,
        "published": published,
    }


# ---------------------------------------------------------------------------
# Excel output
# ---------------------------------------------------------------------------


def write_excel(sheets: dict, out_path: Path) -> None:
    """Write the documentation pack, one sheet per section."""
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        for name, rows in sheets.items():
            frame = pd.DataFrame(rows) if rows else pd.DataFrame({"(nothing found)": []})
            frame.to_excel(writer, sheet_name=name[:31], index=False)

            worksheet = writer.sheets[name[:31]]
            worksheet.freeze_panes = "A2"
            if rows:
                worksheet.auto_filter.ref = worksheet.dimensions

            for column_cells in worksheet.columns:
                letter = column_cells[0].column_letter
                width = max(
                    (len(str(c.value).split("\n")[0]) for c in column_cells if c.value is not None),
                    default=10,
                )
                # Formula and SQL columns get a fixed, readable width rather
                # than one as wide as the longest query.
                header = str(column_cells[0].value or "")
                cap = 60 if header in ("Formula", "Table or Query") else 45
                worksheet.column_dimensions[letter].width = min(max(width + 2, 12), cap)

    print(f"\nDocumentation pack: {out_path}")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------


def document_year(
    year: int = 2026,
    out: str = "./tableau_year",
    date_field: str = "created",
    include_extracts: bool = False,
    overwrite: bool = False,
    dry_run: bool = False,
) -> dict:
    """Download a year's workbooks and build the documentation pack.

    date_field: "created" (first published), "modified", or "either".
    """
    if date_field not in ("created", "modified", "either"):
        raise RuntimeError('date_field must be "created", "modified", or "either".')

    out_dir = Path(out).expanduser().resolve()
    out_dir.mkdir(parents=True, exist_ok=True)
    print(f"Output directory: {out_dir}\n")

    server = connect()

    wb_rows, issues = [], []
    all_connections, all_relations, all_calcs = [], [], []
    all_parameters, all_views, all_published = [], [], []

    try:
        projects = fetch_all(server.projects.get)
        project_paths = build_project_paths(projects)

        users = []
        try:
            users = fetch_all(server.users.get)
        except Exception:  # noqa: BLE001 -- non-admins often can't list users
            pass
        user_names = {u.id: u.name for u in users}

        workbooks = fetch_all(server.workbooks.get)
        print(f"{len(workbooks)} workbooks visible on server")

        def in_year(wb) -> bool:
            created = wb.created_at.year if wb.created_at else None
            modified = wb.updated_at.year if wb.updated_at else None
            if date_field == "created":
                return created == year
            if date_field == "modified":
                return modified == year
            return year in (created, modified)

        selected = [wb for wb in workbooks if in_year(wb)]
        print(f"{len(selected)} published in {year} (by {date_field} date)\n")

        if not selected:
            print(
                "Nothing matched. If workbooks were republished rather than newly\n"
                'created this year, try date_field="either".'
            )
            return {"workbooks": [], "issues": []}

        started = time.time()
        total_bytes = 0

        for i, wb in enumerate(selected, 1):
            parts = project_paths.get(wb.project_id) or [safe_name(wb.project_name)]
            project_display = " > ".join(parts)
            target_dir = out_dir.joinpath(*parts)
            stem = safe_name(wb.name)
            label = f"[{i}/{len(selected)}] {project_display} / {wb.name}"

            row = {
                "Workbook": wb.name,
                "Project Path": project_display,
                "Owner": user_names.get(wb.owner_id, wb.owner_id),
                "Created": fmt_date(wb.created_at),
                "Last Modified": fmt_date(wb.updated_at),
                "Server Size (MB)": mb(wb.size),
                "Tags": ", ".join(sorted(wb.tags)) if wb.tags else "",
                "Worksheets": "",
                "Dashboards": "",
                "Stories": "",
                "Connections": "",
                "Tables / Queries": "",
                "Custom SQL": "",
                "Calculated Fields": "",
                "LOD Expressions": "",
                "Table Calcs": "",
                "Parameters": "",
                "Published Data Sources": "",
                "Blended Worksheets": "",
                "File": "",
                "URL": wb.webpage_url,
            }

            if dry_run:
                row["File"] = str((target_dir / f"{stem}.twb").relative_to(out_dir))
                wb_rows.append(row)
                print(f"{label} -> would download")
                continue

            try:
                existing = next(
                    (p for ext in (".twbx", ".twb") if (p := target_dir / f"{stem}{ext}").exists()),
                    None,
                )
                if existing and not overwrite:
                    package = existing
                    print(f"{label} -> already downloaded")
                else:
                    target_dir.mkdir(parents=True, exist_ok=True)
                    downloaded = Path(
                        server.workbooks.download(
                            wb.id, filepath=str(target_dir), include_extract=include_extracts
                        )
                    )
                    package = unique_path(target_dir, stem, downloaded.suffix)
                    if downloaded != package:
                        downloaded.replace(package)
                    total_bytes += package.stat().st_size
                    print(f"{label} -> {package.name} ({mb(package.stat().st_size)} MB)")

                row["File"] = str(package.relative_to(out_dir))

            except TSC.ServerResponseError as exc:
                issues.append(
                    {
                        "Workbook": wb.name,
                        "Project Path": project_display,
                        "Stage": "Download",
                        "Detail": f"{exc.code}: {exc.summary}",
                    }
                )
                row["File"] = "NOT DOWNLOADED"
                wb_rows.append(row)
                print(f"{label} -> FAILED ({exc.code})")
                continue
            except Exception as exc:  # noqa: BLE001
                issues.append(
                    {
                        "Workbook": wb.name,
                        "Project Path": project_display,
                        "Stage": "Download",
                        "Detail": f"{type(exc).__name__}: {exc}",
                    }
                )
                row["File"] = "NOT DOWNLOADED"
                wb_rows.append(row)
                print(f"{label} -> FAILED ({type(exc).__name__})")
                continue

            # -- Parse the XML -------------------------------------------------
            try:
                parsed = analyze_workbook(open_workbook_xml(package), wb.name)
            except Exception as exc:  # noqa: BLE001
                issues.append(
                    {
                        "Workbook": wb.name,
                        "Project Path": project_display,
                        "Stage": "Parse",
                        "Detail": f"{type(exc).__name__}: {exc}",
                    }
                )
                wb_rows.append(row)
                print("    could not read workbook XML")
                continue

            all_connections += parsed["connections"]
            all_relations += parsed["relations"]
            all_calcs += parsed["calcs"]
            all_parameters += parsed["parameters"]
            all_views += parsed["views"]
            all_published += parsed["published"]

            worksheets = [v for v in parsed["views"] if v["Type"] == "Worksheet"]
            row.update(
                {
                    "Worksheets": len(worksheets),
                    "Dashboards": sum(1 for v in parsed["views"] if v["Type"] == "Dashboard"),
                    "Stories": sum(1 for v in parsed["views"] if v["Type"] == "Story"),
                    "Connections": len(parsed["connections"]),
                    "Tables / Queries": len(parsed["relations"]),
                    "Custom SQL": sum(
                        1 for r in parsed["relations"] if r["Object Type"] == "Custom SQL"
                    ),
                    "Calculated Fields": len(parsed["calcs"]),
                    "LOD Expressions": sum(1 for c in parsed["calcs"] if c["Is LOD"]),
                    "Table Calcs": sum(1 for c in parsed["calcs"] if c["Is Table Calc"]),
                    "Parameters": len(parsed["parameters"]),
                    "Published Data Sources": len(parsed["published"]),
                    "Blended Worksheets": sum(1 for v in worksheets if v["Is Blend"]),
                }
            )
            wb_rows.append(row)

        # -- Summary sheet ------------------------------------------------------
        def total(key: str) -> int:
            return sum(r[key] for r in wb_rows if isinstance(r.get(key), int))

        connection_types = sorted({c["Connection Type"] for c in all_connections})
        summary = [
            {"Metric": "Year", "Value": year},
            {"Metric": "Date basis", "Value": date_field},
            {"Metric": "Workbooks", "Value": len(wb_rows)},
            {"Metric": "Worksheets", "Value": total("Worksheets")},
            {"Metric": "Dashboards", "Value": total("Dashboards")},
            {"Metric": "Stories", "Value": total("Stories")},
            {"Metric": "Database connections", "Value": len(all_connections)},
            {"Metric": "Connection types", "Value": ", ".join(connection_types)},
            {"Metric": "Tables and queries", "Value": len(all_relations)},
            {"Metric": "Custom SQL queries", "Value": total("Custom SQL")},
            {"Metric": "Calculated fields", "Value": len(all_calcs)},
            {"Metric": "LOD expressions", "Value": total("LOD Expressions")},
            {"Metric": "Table calculations", "Value": total("Table Calcs")},
            {"Metric": "Parameters", "Value": len(all_parameters)},
            {"Metric": "Blended worksheets", "Value": total("Blended Worksheets")},
            {
                "Metric": "Published data source links",
                "Value": len(all_published),
            },
            {"Metric": "Extracts included in download", "Value": include_extracts},
            {"Metric": "Issues", "Value": len(issues)},
        ]

        excel_path = out_dir / f"Tableau_{year}_Report_Documentation.xlsx"
        write_excel(
            {
                "Summary": summary,
                "Workbooks": wb_rows,
                "Connections": all_connections,
                "Tables & Queries": all_relations,
                "Calculated Fields": all_calcs,
                "Parameters": all_parameters,
                "Views": all_views,
                "Published Data Srcs": all_published,
                "Issues": issues,
            },
            excel_path,
        )

        if issues:
            issues_csv = out_dir / "_issues.csv"
            with issues_csv.open("w", newline="", encoding="utf-8-sig") as handle:
                writer = csv.DictWriter(handle, fieldnames=list(issues[0].keys()))
                writer.writeheader()
                writer.writerows(issues)
            print(f"Issues:             {issues_csv}")

        print(f"\n{'-' * 60}")
        print(f"Finished in {(time.time() - started) / 60:.1f} min")
        print(f"  {'Workbooks documented':<28} {len(wb_rows)}")
        print(f"  {'Custom SQL queries':<28} {total('Custom SQL')}")
        print(f"  {'Calculated fields':<28} {len(all_calcs)}")
        print(f"  {'Connection types':<28} {', '.join(connection_types) or 'none found'}")
        if total_bytes:
            print(f"  {'Downloaded':<28} {total_bytes / 1_073_741_824:.2f} GB")
        if issues:
            print(f"  {'Issues':<28} {len(issues)}")

        return {"workbooks": wb_rows, "issues": issues}

    finally:
        try:
            server.auth.sign_out()
            print("Signed out.")
        except Exception:  # noqa: BLE001
            pass


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------


def _in_notebook() -> bool:
    try:
        from IPython import get_ipython

        return get_ipython() is not None
    except ImportError:
        return False


def _parse_cli_args():
    import argparse

    parser = argparse.ArgumentParser(
        description="Download and document a year's Tableau workbooks",
        allow_abbrev=False,
    )
    parser.add_argument("--year", type=int, default=2026)
    parser.add_argument("--out", default="./tableau_year")
    parser.add_argument(
        "--date-field", default="created", choices=["created", "modified", "either"]
    )
    parser.add_argument("--include-extracts", action="store_true")
    parser.add_argument("--overwrite", action="store_true")
    parser.add_argument("--dry-run", action="store_true")
    args, _unknown = parser.parse_known_args()
    return args


if __name__ == "__main__" and not _in_notebook():
    cli = _parse_cli_args()
    try:
        document_year(
            year=cli.year,
            out=cli.out,
            date_field=cli.date_field,
            include_extracts=cli.include_extracts,
            overwrite=cli.overwrite,
            dry_run=cli.dry_run,
        )
    except RuntimeError as err:
        print(f"\n{err}", file=sys.stderr)
