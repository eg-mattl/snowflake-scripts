from __future__ import annotations

import tempfile
import unittest
from datetime import datetime
from pathlib import Path
from unittest.mock import patch

from openpyxl import load_workbook

import export_user_default_roles as exporter


class FakeCursor:
    def __init__(self) -> None:
        self.description = None
        self._rows = []
        self.executed_sql = ""

    def __enter__(self) -> "FakeCursor":
        return self

    def __exit__(self, exc_type, exc_value, traceback) -> None:
        return None

    def execute(self, sql: str) -> None:
        self.executed_sql = sql
        if "CURRENT_ACCOUNT()" in sql:
            self.description = [
                ("ACCOUNT",),
                ("SESSION_USER",),
                ("SESSION_ROLE",),
                ("SESSION_WAREHOUSE",),
            ]
            self._rows = [("TEST_ACCOUNT", "REPORT_USER", "READ_ONLY", "COMPUTE_WH")]
        else:
            self.description = [
                ("USER_NAME",),
                ("LOGIN_NAME",),
                ("DISPLAY_NAME",),
                ("USER_TYPE",),
                ("DISABLED",),
                ("DEFAULT_ROLE",),
            ]
            self._rows = [
                ("ALICE", "alice", "Alice Example", "PERSON", False, "ANALYST"),
                ("SVC_EXPORT", "svc_export", "Export Service", "SERVICE", True, None),
            ]

    def fetchone(self):
        return self._rows[0] if self._rows else None

    def fetchall(self):
        return self._rows


class FakeConnection:
    def cursor(self) -> FakeCursor:
        return FakeCursor()


class ExportUserDefaultRolesTests(unittest.TestCase):
    def test_connection_parameters_use_browser_sso_without_secrets(self) -> None:
        env = {
            "SNOWFLAKE_ACCOUNT": "test-account",
            "SNOWFLAKE_USER": "test-user",
            "SNOWFLAKE_ROLE": "READ_ONLY",
            "SNOWFLAKE_WAREHOUSE": "COMPUTE_WH",
        }
        params = exporter.build_connection_parameters(env)
        self.assertEqual(params["authenticator"], "externalbrowser")
        self.assertEqual(params["account"], "test-account")
        self.assertEqual(params["role"], "READ_ONLY")
        self.assertNotIn("password", params)

    def test_missing_connection_settings_are_reported(self) -> None:
        with self.assertRaisesRegex(ValueError, "SNOWFLAKE_ACCOUNT, SNOWFLAKE_USER"):
            exporter.build_connection_parameters({})

    def test_jwt_requires_private_key_path(self) -> None:
        env = {
            "SNOWFLAKE_ACCOUNT": "test-account",
            "SNOWFLAKE_USER": "test-user",
            "SNOWFLAKE_AUTHENTICATOR": "SNOWFLAKE_JWT",
        }
        with self.assertRaisesRegex(ValueError, "SNOWFLAKE_PRIVATE_KEY_FILE"):
            exporter.build_connection_parameters(env)

    def test_fetch_uses_all_non_deleted_account_users(self) -> None:
        context, users = exporter.fetch_report_data(FakeConnection())
        self.assertEqual(context["ACCOUNT"], "TEST_ACCOUNT")
        self.assertEqual(len(users), 2)
        self.assertEqual(users[1]["USER_TYPE"], "SERVICE")
        self.assertTrue(users[1]["DISABLED"])
        self.assertIsNone(users[1]["DEFAULT_ROLE"])
        self.assertIn("SNOWFLAKE.ACCOUNT_USAGE.USERS", exporter.USERS_QUERY)
        self.assertIn("WHERE DELETED_ON IS NULL", exporter.USERS_QUERY)
        self.assertNotIn("DISABLED = FALSE", exporter.USERS_QUERY)

    def test_workbook_has_summary_and_literal_not_set_role(self) -> None:
        context, users = exporter.fetch_report_data(FakeConnection())
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "reports" / "user_roles.xlsx"
            generated_at = datetime(2026, 11, 9, 12, 0, 0)
            result = exporter.write_workbook(
                output,
                context,
                users,
                generated_at=generated_at,
            )

            self.assertTrue(result.is_file())
            workbook = load_workbook(result, data_only=False)
            self.assertEqual(workbook.sheetnames, ["Summary", "Users"])
            summary = workbook["Summary"]
            self.assertEqual(summary["B9"].value, 2)
            self.assertEqual(summary["B10"].value, 1)
            self.assertEqual(summary["B11"].value, 1)

            detail = workbook["Users"]
            self.assertEqual(detail.freeze_panes, "A2")
            self.assertIn("SnowflakeUsers", detail.tables)
            self.assertEqual(detail["F2"].value, "ANALYST")
            self.assertEqual(detail["F3"].value, "(not set)")
            self.assertEqual(detail["E3"].value, True)
            self.assertEqual(detail["A2"].data_type, "s")

    def test_metadata_that_looks_like_a_formula_is_saved_as_text(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "formula-safety.xlsx"
            user = {
                "USER_NAME": "=HYPERLINK(\"https://example.com\",\"click\")",
                "LOGIN_NAME": "formula_test",
                "DISPLAY_NAME": "Formula Test",
                "USER_TYPE": "PERSON",
                "DISABLED": False,
                "DEFAULT_ROLE": "=1+1",
            }
            exporter.write_workbook(output, {}, [user])
            detail = load_workbook(output, data_only=False)["Users"]
            self.assertEqual(detail["A2"].value, user["USER_NAME"])
            self.assertEqual(detail["A2"].data_type, "s")
            self.assertEqual(detail["F2"].value, "=1+1")
            self.assertEqual(detail["F2"].data_type, "s")

    def test_existing_output_requires_force(self) -> None:
        context, users = exporter.fetch_report_data(FakeConnection())
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "report.xlsx"
            output.write_bytes(b"existing file")
            with self.assertRaises(FileExistsError):
                exporter.write_workbook(output, context, users)
            self.assertEqual(output.read_bytes(), b"existing file")

    def test_connection_environment_is_read_at_call_time(self) -> None:
        with patch.dict(
            "os.environ",
            {"SNOWFLAKE_ACCOUNT": "runtime-account", "SNOWFLAKE_USER": "runtime-user"},
            clear=True,
        ):
            params = exporter.build_connection_parameters()
        self.assertEqual(params["account"], "runtime-account")
        self.assertEqual(params["user"], "runtime-user")


if __name__ == "__main__":
    unittest.main()
