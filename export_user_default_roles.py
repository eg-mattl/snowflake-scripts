#!/usr/bin/env python3
"""Export non-deleted Snowflake users and their configured default roles to Excel."""

from __future__ import annotations

import argparse
import os
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Mapping, Sequence

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo


CONTEXT_QUERY = """
SELECT
    CURRENT_ACCOUNT() AS ACCOUNT,
    CURRENT_USER() AS SESSION_USER,
    CURRENT_ROLE() AS SESSION_ROLE,
    CURRENT_WAREHOUSE() AS SESSION_WAREHOUSE
"""

# ACCOUNT_USAGE retains service and disabled accounts. Filter only deleted users
# so the report reflects every non-deleted account user, not just human users.
USERS_QUERY = """
SELECT
    NAME AS USER_NAME,
    LOGIN_NAME,
    DISPLAY_NAME,
    TYPE AS USER_TYPE,
    DISABLED,
    DEFAULT_ROLE
FROM SNOWFLAKE.ACCOUNT_USAGE.USERS
WHERE DELETED_ON IS NULL
ORDER BY NAME
"""

USER_HEADERS = (
    "User Name",
    "Login Name",
    "Display Name",
    "User Type",
    "Disabled",
    "Default Role",
)

HEADER_FILL = PatternFill(fill_type="solid", fgColor="17365D")
HEADER_FONT = Font(color="FFFFFF", bold=True)


def build_connection_parameters(env: Mapping[str, str] | None = None) -> dict[str, str]:
    """Build Snowflake Connector parameters from environment, never from CLI secrets."""
    source = os.environ if env is None else env
    required = ("SNOWFLAKE_ACCOUNT", "SNOWFLAKE_USER")
    missing = [name for name in required if not source.get(name)]
    if missing:
        raise ValueError(
            "Missing required environment variable(s): " + ", ".join(missing)
        )

    authenticator = source.get("SNOWFLAKE_AUTHENTICATOR", "externalbrowser")
    params = {
        "account": source["SNOWFLAKE_ACCOUNT"],
        "user": source["SNOWFLAKE_USER"],
        "authenticator": authenticator,
    }
    for env_name, connector_name in (
        ("SNOWFLAKE_ROLE", "role"),
        ("SNOWFLAKE_WAREHOUSE", "warehouse"),
    ):
        if source.get(env_name):
            params[connector_name] = source[env_name]

    if authenticator.upper() == "SNOWFLAKE_JWT":
        key_file = source.get("SNOWFLAKE_PRIVATE_KEY_FILE")
        if not key_file:
            raise ValueError(
                "SNOWFLAKE_PRIVATE_KEY_FILE is required when "
                "SNOWFLAKE_AUTHENTICATOR=SNOWFLAKE_JWT"
            )
        params["private_key_file"] = str(Path(key_file).expanduser())
        passphrase = source.get("SNOWFLAKE_PRIVATE_KEY_FILE_PWD")
        if passphrase:
            params["private_key_file_pwd"] = passphrase

    return params


def _column_names(description: Sequence[Any]) -> list[str]:
    """Support DB-API tuple descriptions and newer connector metadata objects."""
    names: list[str] = []
    for column in description:
        name = getattr(column, "name", None)
        if name is None:
            name = column[0]
        names.append(str(name))
    return names


def fetch_report_data(connection: Any) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    """Fetch session context and the complete non-deleted user inventory."""
    with connection.cursor() as cursor:
        cursor.execute(CONTEXT_QUERY)
        context_row = cursor.fetchone()
        context_names = _column_names(cursor.description or ())
        context = dict(zip(context_names, context_row or ()))

    with connection.cursor() as cursor:
        cursor.execute(USERS_QUERY)
        user_names = _column_names(cursor.description or ())
        users = [dict(zip(user_names, row)) for row in cursor.fetchall()]

    return context, users


def _has_default_role(user: Mapping[str, Any]) -> bool:
    role = user.get("DEFAULT_ROLE")
    return role is not None and bool(str(role).strip())


def _write_cell(cell: Any, value: Any) -> None:
    """Write metadata as literal text so names cannot become Excel formulas."""
    if value is None:
        cell.value = ""
    elif isinstance(value, (datetime, bool, int, float)):
        cell.value = value
    else:
        cell.value = str(value)
        cell.data_type = "s"


def _style_header(sheet: Any) -> None:
    for cell in sheet[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(vertical="center")
    sheet.row_dimensions[1].height = 22


def _fit_columns(sheet: Any, max_width: int = 60) -> None:
    for column_cells in sheet.columns:
        letter = column_cells[0].column_letter
        longest = max(
            (len(str(cell.value)) for cell in column_cells if cell.value is not None),
            default=0,
        )
        sheet.column_dimensions[letter].width = min(max(longest + 2, 12), max_width)


def write_workbook(
    output_path: str | Path,
    context: Mapping[str, Any],
    users: Sequence[Mapping[str, Any]],
    *,
    force: bool = False,
    generated_at: datetime | None = None,
) -> Path:
    """Create a Summary + Users workbook. Existing files require --force."""
    path = Path(output_path).expanduser().resolve()
    if path.exists() and not force:
        raise FileExistsError(f"Output already exists; pass --force to replace it: {path}")
    path.parent.mkdir(parents=True, exist_ok=True)

    generated_at = generated_at or datetime.now(timezone.utc).replace(tzinfo=None)
    users_with_role = sum(1 for user in users if _has_default_role(user))
    users_without_role = len(users) - users_with_role

    workbook = Workbook()
    workbook.properties.title = "Snowflake User Default Roles"
    workbook.properties.subject = "Non-deleted Snowflake users and configured default roles"
    workbook.properties.creator = "Data Engineering"

    summary = workbook.active
    if summary is None:
        raise RuntimeError("Could not initialize the workbook summary worksheet.")
    summary.title = "Summary"
    summary.append(("Metric", "Value"))
    summary_rows = (
        ("Report generated (UTC)", generated_at),
        ("Snowflake account", context.get("ACCOUNT") or "(unavailable)"),
        ("Session user", context.get("SESSION_USER") or "(unavailable)"),
        ("Session role", context.get("SESSION_ROLE") or "(unavailable)"),
        ("Session warehouse", context.get("SESSION_WAREHOUSE") or "(none)"),
        ("Source", "SNOWFLAKE.ACCOUNT_USAGE.USERS"),
        ("Scope", "All rows where DELETED_ON IS NULL; includes disabled and service users."),
        ("Non-deleted user rows", len(users)),
        ("Users with a default role", users_with_role),
        ("Users without a default role", users_without_role),
        ("Null default role display", "(not set)"),
        (
            "Freshness note",
            "ACCOUNT_USAGE metadata can lag recent account changes by about two hours.",
        ),
    )
    for label, value in summary_rows:
        row = summary.max_row + 1
        summary.cell(row=row, column=1, value=label)
        _write_cell(summary.cell(row=row, column=2), value)
    _style_header(summary)
    summary.freeze_panes = "A2"
    summary.auto_filter.ref = f"A1:B{summary.max_row}"
    _fit_columns(summary)
    summary.column_dimensions["A"].width = 34
    summary.column_dimensions["B"].width = 88
    summary["B2"].number_format = "yyyy-mm-dd hh:mm:ss"
    summary["B2"].alignment = Alignment(vertical="top")
    for row in summary.iter_rows(min_row=2, min_col=2, max_col=2):
        row[0].alignment = Alignment(vertical="top", wrap_text=True)

    detail = workbook.create_sheet("Users")
    detail.append(USER_HEADERS)
    for user in users:
        values = (
            user.get("USER_NAME"),
            user.get("LOGIN_NAME"),
            user.get("DISPLAY_NAME"),
            user.get("USER_TYPE"),
            user.get("DISABLED"),
            user.get("DEFAULT_ROLE") or "(not set)",
        )
        detail.append([None] * len(USER_HEADERS))
        row_number = detail.max_row
        for column_number, value in enumerate(values, start=1):
            _write_cell(detail.cell(row=row_number, column=column_number), value)

    _style_header(detail)
    detail.freeze_panes = "A2"
    if users:
        table = Table(displayName="SnowflakeUsers", ref=f"A1:F{detail.max_row}")
        table.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
        detail.add_table(table)
    else:
        detail.auto_filter.ref = "A1:F1"
    _fit_columns(detail)

    workbook.save(path)
    return path


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Export all non-deleted Snowflake users and their configured default "
            "roles to an Excel workbook."
        )
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=Path("snowflake_user_default_roles.xlsx"),
        help="Output .xlsx path (default: ./snowflake_user_default_roles.xlsx)",
    )
    parser.add_argument(
        "--force",
        action="store_true",
        help="Replace the output file if it already exists.",
    )
    args = parser.parse_args(argv)

    try:
        connection_parameters = build_connection_parameters()
        import snowflake.connector

        connection = snowflake.connector.connect(**connection_parameters)
        try:
            context, users = fetch_report_data(connection)
        finally:
            connection.close()

        output_path = write_workbook(
            args.output,
            context,
            users,
            force=args.force,
        )
    except Exception as exc:  # Show connector/file errors without a traceback by default.
        print(f"Export failed: {exc}", file=sys.stderr)
        return 1

    with_role = sum(1 for user in users if _has_default_role(user))
    print(f"Created: {output_path}")
    print(f"Non-deleted user rows: {len(users)}")
    print(f"Users with a default role: {with_role}")
    print(f"Users without a default role: {len(users) - with_role}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
