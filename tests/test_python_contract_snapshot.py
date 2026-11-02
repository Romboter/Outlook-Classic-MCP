"""Offline regression checks for the Python-to-Rust compatibility baseline.

Does not import Outlook, pythoncom, win32com, FastMCP, or the COM bridge.
"""
from __future__ import annotations

import json
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
from verify_python_contract_snapshot import compare_static  # noqa: E402

DECLARATIONS = ROOT / "rust/compatibility/python-tools.declarations.json"
RESULTS = ROOT / "rust/compatibility/python-results.redacted.json"
RUNTIME = ROOT / "rust/compatibility/python-tools.runtime.json"


class PythonContractSnapshotTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.declarations = json.loads(DECLARATIONS.read_text(encoding="utf-8"))
        cls.results = json.loads(RESULTS.read_text(encoding="utf-8"))
        cls.runtime = json.loads(RUNTIME.read_text(encoding="utf-8"))
        cls.tools = {tool["name"]: tool for tool in cls.declarations["tools"]}

    def test_all_thirty_one_tools_have_a_single_input_and_result_entry(self):
        names = set(self.tools)
        self.assertEqual(len(names), 31)
        self.assertEqual(set(item["name"] for item in self.results["tools"]), names)
        self.assertEqual(self.declarations["toolCount"], 31)
        self.assertEqual(self.results["toolCount"], 31)

    def test_captured_fastmcp_registry_covers_the_same_31_tools(self):
        self.assertEqual(self.runtime["toolCount"], 31)
        actual = {tool["name"]: tool for tool in self.runtime["tools"]}
        self.assertEqual(set(actual), set(self.tools))
        for name, tool in actual.items():
            self.assertEqual(
                set(tool["inputSchema"]["properties"]),
                set(self.tools[name]["inputSchema"]["properties"]),
            )

    def test_static_schema_stays_in_sync_with_python_sources(self):
        self.assertEqual(compare_static(self.declarations), [])

    def test_python_defaults_are_preserved_without_imposing_rust_limits(self):
        list_mail = self.tools["outlook_list_mails"]["inputSchema"]
        self.assertEqual(list_mail["properties"]["limit"]["default"], 25)
        self.assertEqual(list_mail["properties"]["limit"]["maximum"], 100)
        self.assertEqual(list_mail["properties"]["folder"]["default"], "inbox")
        self.assertEqual(self.tools["outlook_search_mails"]["inputSchema"]["required"], ["query"])
        self.assertEqual(
            self.tools["outlook_search_mails"]["inputSchema"]["properties"]["scope"]["default"],
            "subject_body",
        )
        self.assertEqual(
            self.tools["outlook_list_events"]["inputSchema"]["properties"]["limit"]["maximum"],
            200,
        )

    def test_nested_recurrence_contract_is_recorded(self):
        schema = self.tools["outlook_create_event"]["inputSchema"]
        self.assertEqual(schema["required"], ["subject", "start", "end"])
        self.assertIn("Recurrence", schema["$defs"])
        self.assertEqual(
            schema["$defs"]["Recurrence"]["properties"]["type"]["enum"],
            ["daily", "weekly", "monthly", "yearly"],
        )

    def test_annotations_and_five_ui_resources(self):
        ui = [t for t in self.tools.values() if t["uiResource"]]
        self.assertEqual(len(ui), 7)  # mail list/search/get, calendar, event detail, contacts and tasks
        self.assertFalse(self.tools["outlook_send_mail"]["annotations"]["readOnlyHint"])
        self.assertFalse(self.tools["outlook_send_mail"]["annotations"]["idempotentHint"])
        self.assertTrue(self.tools["outlook_delete_mail"]["annotations"]["destructiveHint"])
        self.assertFalse(self.tools["outlook_toggle_rule"]["annotations"]["readOnlyHint"])
        self.assertTrue(self.tools["outlook_whoami"]["annotations"]["readOnlyHint"])

    def test_fixtures_contain_only_explicit_placeholders(self):
        sensitive_keys = ("entry_id", "new_entry_id", "in_reply_to", "forwarded",
                          "subject", "body", "email", "address", "query", "name")
        for item in self.results["tools"]:
            self.assertFalse(item["verifiedAgainstLiveOutlook"])
            self.assertIsInstance(item["redactedFixture"], dict)
            fixture_keys = set(item["redactedFixture"])
            self.assertTrue(fixture_keys <= set(item["topLevelKeys"]))
            conditional = {
                "outlook_get_mail": {"body_truncated", "body_total_chars", "html_body"},
                "outlook_send_mail": {"entry_id"},
                "outlook_resolve_name": {"note"},
                "outlook_list_mails": {"next_offset"},
            }.get(item["name"], set())
            self.assertEqual(set(item["topLevelKeys"]) - fixture_keys, conditional)
            for key, value in item["redactedFixture"].items():
                if any(term in key for term in sensitive_keys):
                    if isinstance(value, str):
                        self.assertEqual(value, "[REDACTED]")
                    elif isinstance(value, list):
                        self.assertEqual(value, [])


if __name__ == "__main__":
    unittest.main()
