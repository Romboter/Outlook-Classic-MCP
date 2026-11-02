"""Exercise existing Python FastMCP wrappers with a synthetic bridge, never COM.

These tests validate wrapping/formatting and preserve Python's odd
structuredContent envelope. They do NOT test Outlook, STA or transport stdio.
Run only in the Windows contract-CI job with project dependencies installed.
"""
from __future__ import annotations

import asyncio
import json
import unittest

from mcp.server.fastmcp import FastMCP
from mcp.types import CallToolResult
from outlook_mcp.tools import register_all


class SyntheticBridge:
    async def call(self, fn, *args, **kwargs):
        name = fn.__name__
        if name == "whoami":
            return {
                "current_user": "Synthetic User",
                "accounts": [],
                "local_time": "2026-01-01T00:00:00+00:00",
                "timezone": "UTC",
                "utc_offset": "+0000",
            }
        if name == "list_folders":
            return [{"name": "Example", "path": "Example",
                     "item_count": 1, "unread_count": 0,
                     "default_item_type": 0}]
        if name == "list_mails":
            return {
                "folder": "Example", "count": 1, "offset": 0,
                "limit": kwargs["limit"], "has_more": False,
                "items": [{"entry_id": "SYNTHETIC_ONLY",
                           "subject": "Synthetic subject",
                           "from": "Synthetic User", "from_address": "",
                           "received": "2026-01-01T00:00:00+00:00",
                           "unread": False, "has_attachments": False}],
            }
        if name == "get_mail":
            full = "A" * 280
            budget = kwargs["max_body_chars"]
            payload = {"entry_id": "SYNTHETIC_ONLY", "subject": "Synthetic",
                       "body": full[:budget] if budget else full}
            if budget and len(full) > budget:
                payload["body_truncated"] = True
                payload["body_total_chars"] = len(full)
            return payload
        raise AssertionError("Unexpected tool attempted through synthetic bridge")


def unwrap_result(value):
    # FastMCP v1 tool conversion may return (text blocks, structured data),
    # while an explicitly returned CallToolResult is already an MCP object.
    if isinstance(value, CallToolResult):
        return value.content, value.structuredContent
    if isinstance(value, tuple) and len(value) == 2:
        return value
    raise AssertionError(f"Unexpected FastMCP result shape: {type(value).__name__}")


class OfflineWrapperContractTests(unittest.IsolatedAsyncioTestCase):
    async def asyncSetUp(self):
        self.server = FastMCP("synthetic-no-outlook")
        register_all(self.server, SyntheticBridge())

    async def tool(self, name, arguments):
        return unwrap_result(await self.server.call_tool(name, arguments))

    async def test_string_wrappers_keep_result_envelope(self):
        for name in ("outlook_whoami", "outlook_list_folders"):
            with self.subTest(tool=name):
                blocks, structured = await self.tool(name, {"response_format": "json"})
                self.assertEqual(set(structured), {"result"})
                self.assertIsInstance(structured["result"], str)
                self.assertEqual(json.loads(blocks[0].text), json.loads(structured["result"]))
                if name == "outlook_list_folders":
                    self.assertEqual(set(json.loads(blocks[0].text)), {"count", "items"})

    async def test_call_tool_result_wrappers_keep_direct_payload(self):
        blocks, structured = await self.tool(
            "outlook_list_mails", {"response_format": "json"}
        )
        self.assertEqual(json.loads(blocks[0].text), structured)
        self.assertEqual(structured["count"], 1)
        blocks, structured = await self.tool(
            "outlook_list_mails", {"response_format": "markdown"}
        )
        self.assertIn("Synthetic subject", blocks[0].text)
        self.assertEqual(structured["count"], 1)

    async def test_selected_mail_truncation_fields_are_conditional(self):
        for limit, expected_chars in [(200, 200), (0, 280)]:
            with self.subTest(max_body_chars=limit):
                blocks, structured = await self.tool(
                    "outlook_get_mail",
                    {"entry_id": "SYNTHETIC_ONLY",
                     "response_format": "json",
                     "max_body_chars": limit},
                )
                self.assertEqual(json.loads(blocks[0].text), structured)
                self.assertEqual(len(structured["body"]), expected_chars)
                if limit:
                    self.assertTrue(structured["body_truncated"])
                    self.assertEqual(structured["body_total_chars"], 280)
                else:
                    self.assertNotIn("body_truncated", structured)
                    self.assertNotIn("body_total_chars", structured)


if __name__ == "__main__":
    unittest.main()
