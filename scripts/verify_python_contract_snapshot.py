"""Check the offline Python MCP declaration snapshot without starting Outlook.

Commands (from repository root):
    python scripts/verify_python_contract_snapshot.py --static
    python scripts/verify_python_contract_snapshot.py --runtime
    python scripts/verify_python_contract_snapshot.py --capture-runtime

--static uses only Python's standard library and reads source syntax.
--runtime and --capture-runtime need the project's Python dependencies, but
NEVER start the Outlook COM bridge and NEVER call a tool. They inspect only
FastMCP's locally registered metadata. Runtime output has no mailbox data.
"""
from __future__ import annotations

import argparse
import ast
import asyncio
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SNAPSHOT = ROOT / "rust/compatibility/python-tools.declarations.json"
RUNTIME = ROOT / "rust/compatibility/python-tools.runtime.json"
TOOLS_DIR = ROOT / "src/outlook_mcp/tools"
EXPECTED = 31


def get_tool_decorator(fn: ast.AsyncFunctionDef) -> ast.Call | None:
    for node in fn.decorator_list:
        if (
            isinstance(node, ast.Call)
            and isinstance(node.func, ast.Attribute)
            and node.func.attr == "tool"
            and isinstance(node.func.value, ast.Name)
            and node.func.value.id == "mcp"
        ):
            return node
    return None


def get_keyword(node: ast.Call, name: str) -> ast.expr | None:
    return next((kw.value for kw in node.keywords if kw.arg == name), None)


def as_literal(node: ast.expr | None):
    if node is None:
        return None
    return ast.literal_eval(node)


def type_spec(node: ast.expr) -> dict:
    if isinstance(node, ast.Name):
        return {
            "str": {"type": "string"},
            "int": {"type": "integer"},
            "bool": {"type": "boolean"},
            "Recurrence": {"$ref": "#/$defs/Recurrence"},
        }[node.id]
    if isinstance(node, ast.Subscript) and isinstance(node.value, ast.Name):
        if node.value.id == "Optional":
            return {"anyOf": [type_spec(node.slice), {"type": "null"}]}
        if node.value.id == "list":
            return {"type": "array", "items": type_spec(node.slice)}
    raise ValueError(f"Unrecognized Python tool parameter type: {ast.unparse(node)}")


def get_parameters(fn: ast.AsyncFunctionDef) -> tuple[dict, list[str]]:
    args = fn.args.args
    defaults = [None] * (len(args) - len(fn.args.defaults)) + list(fn.args.defaults)
    result: dict = {}
    required: list[str] = []
    for arg, default in zip(args, defaults):
        if arg.annotation is None:
            raise ValueError(f"Unannotated parameter: {fn.name}.{arg.arg}")
        annotation = arg.annotation
        field = None
        if (
            isinstance(annotation, ast.Subscript)
            and isinstance(annotation.value, ast.Name)
            and annotation.value.id == "Annotated"
            and isinstance(annotation.slice, ast.Tuple)
        ):
            annotation, *metadata = annotation.slice.elts
            field = next(
                (
                    item
                    for item in metadata
                    if isinstance(item, ast.Call)
                    and isinstance(item.func, ast.Name)
                    and item.func.id == "Field"
                ),
                None,
            )
        schema = type_spec(annotation)
        if default is None:
            required.append(arg.arg)
        else:
            schema["default"] = as_literal(default)
        if field:
            for kw in field.keywords:
                if kw.arg == "description":
                    schema["description"] = as_literal(kw.value)
                elif kw.arg in ("ge", "gt", "le", "lt", "min_length", "max_length"):
                    key = {
                        "ge": "minimum",
                        "gt": "exclusiveMinimum",
                        "le": "maximum",
                        "lt": "exclusiveMaximum",
                        "min_length": "minItems" if schema.get("type") == "array" else "minLength",
                        "max_length": "maxItems" if schema.get("type") == "array" else "maxLength",
                    }[kw.arg]
                    schema[key] = as_literal(kw.value)
        result[arg.arg] = schema
    return result, required


def source_tools() -> dict[str, dict]:
    tools = {}
    for path in sorted(TOOLS_DIR.glob("*.py")):
        module = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
        for node in ast.walk(module):
            if not isinstance(node, ast.AsyncFunctionDef):
                continue
            decorator = get_tool_decorator(node)
            if decorator is None:
                continue
            tool_name = as_literal(get_keyword(decorator, "name")) or node.name
            if tool_name in tools:
                raise ValueError(f"Duplicate Python tool: {tool_name}")
            props, required = get_parameters(node)
            annotation = get_keyword(decorator, "annotations")
            tool_annotation = as_literal(annotation) if annotation is not None else {}
            meta = get_keyword(decorator, "meta")
            view = None
            if (
                isinstance(meta, ast.Call)
                and isinstance(meta.func, ast.Name)
                and meta.func.id == "ui_meta"
                and meta.args
            ):
                view = as_literal(meta.args[0])
            tools[tool_name] = {
                "name": tool_name,
                "source": path.relative_to(ROOT).as_posix(),
                "properties": props,
                "required": required,
                "annotations": tool_annotation,
                "uiResource": f"ui://outlook/{view}.html" if view else None,
                "returnAnnotation": ast.unparse(node.returns) if node.returns else None,
            }
    if len(tools) != EXPECTED:
        raise ValueError(f"Expected {EXPECTED} tools, found {len(tools)}")
    return tools


def compare_static(snapshot: dict) -> list[str]:
    errors = []
    declared = snapshot.get("tools", [])
    by_name = {t["name"]: t for t in declared}
    source = source_tools()
    if snapshot.get("toolCount") != EXPECTED or len(by_name) != EXPECTED:
        errors.append("Snapshot must contain exactly 31 distinct tools")
    for name in sorted(source.keys() | by_name.keys()):
        if name not in source or name not in by_name:
            errors.append(f"{name}: missing from {'source' if name not in source else 'snapshot'}")
            continue
        a, b = source[name], by_name[name]
        input_schema = b["inputSchema"]
        for key, left, right in [
            ("source path", a["source"], b["source"]),
            ("parameters", a["properties"], input_schema["properties"]),
            ("required", a["required"], input_schema["required"]),
            ("annotations", a["annotations"], b["annotations"]),
            ("UI resource", a["uiResource"], b["uiResource"]),
            ("return annotation", a["returnAnnotation"], b["returnContract"]["returnAnnotation"]),
        ]:
            if left != right:
                errors.append(f"{name}: {key} differs from Python source")
    return errors


def normalized_runtime_tool(tool) -> dict:
    # Tool is a Pydantic MCP protocol model. Explicit by_alias retains inputSchema.
    data = tool.model_dump(mode="json", by_alias=True, exclude_none=True)
    return data


async def runtime_tools():
    # build_server only creates a bridge object. Crucially, do not call
    # bridge.start(), mcp.run(), tool functions, or the MCP client.
    sys.path.insert(0, str(ROOT / "src"))
    from outlook_mcp.server import build_server

    server, _bridge = build_server()
    items = await server.list_tools()
    data = [normalized_runtime_tool(item) for item in items]
    return sorted(data, key=lambda item: item["name"])


def compare_runtime(snapshot: dict, actual: list[dict]) -> list[str]:
    errors = []
    reference = {item["name"]: item for item in snapshot["tools"]}
    live = {item["name"]: item for item in actual}
    if set(reference) != set(live):
        errors.append(f"tools/list names differ; missing={sorted(set(reference)-set(live))}, "
                      f"extra={sorted(set(live)-set(reference))}")
    for name in sorted(set(reference) & set(live)):
        a = reference[name]
        b = live[name]
        schema = b.get("inputSchema", {})
        if set(schema.get("properties", {})) != set(a["inputSchema"]["properties"]):
            errors.append(f"{name}: registered property names differ")
            continue
        if set(schema.get("required", [])) != set(a["inputSchema"]["required"]):
            errors.append(f"{name}: registered required fields differ")
        for prop, expected in a["inputSchema"]["properties"].items():
            actual_prop = schema["properties"][prop]
            for key in ("default", "description", "minimum", "maximum",
                        "minLength", "maxLength", "minItems", "maxItems"):
                if key in expected:
                    actual_value = actual_prop.get(key)
                    # Pydantic places Optional[int] constraints on the
                    # non-null anyOf branch; source snapshot keeps them flat.
                    if actual_value is None and "anyOf" in actual_prop:
                        actual_value = next(
                            (part[key] for part in actual_prop["anyOf"] if key in part),
                            None,
                        )
                    if expected[key] != actual_value:
                        errors.append(f"{name}.{prop}: {key} differs")
            # FastMCP can include title and Pydantic-specific schema fields.
            # Compare fundamental types and Optional-nullability only.
            expected_types = sorted(s.get("type", "$ref") for s in
                                    expected.get("anyOf", [expected]))
            actual_types = sorted(s.get("type", "$ref") for s in
                                  actual_prop.get("anyOf", [actual_prop]))
            if expected_types != actual_types:
                errors.append(f"{name}.{prop}: registered type differs")
        annotations = b.get("annotations", {})
        for key, expected in a["annotations"].items():
            if annotations.get(key) != expected:
                errors.append(f"{name}: annotation {key} differs")
        if a["uiResource"] is not None:
            uri = (b.get("_meta") or {}).get("ui", {}).get("resourceUri")
            if uri != a["uiResource"]:
                errors.append(f"{name}: UI resource differs")
    return errors


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--static", action="store_true", help="Check source declarations (no dependencies)")
    group.add_argument("--runtime", action="store_true", help="Compare local FastMCP registration (no Outlook)")
    group.add_argument("--capture-runtime", action="store_true", help="Write actual FastMCP metadata (no Outlook)")
    args = parser.parse_args()
    snapshot = json.loads(SNAPSHOT.read_text(encoding="utf-8"))
    errors = compare_static(snapshot)
    if not args.static:
        runtime = asyncio.run(runtime_tools())
        if len(runtime) != EXPECTED:
            errors.append(f"FastMCP registered {len(runtime)} tools, expected 31")
        errors.extend(compare_runtime(snapshot, runtime))
        if args.runtime and RUNTIME.exists():
            frozen = json.loads(RUNTIME.read_text(encoding="utf-8"))
            if frozen.get("toolCount") != EXPECTED or frozen.get("tools") != runtime:
                errors.append(
                    "FastMCP metadata differs from committed python-tools.runtime.json; "
                    "review the actual metadata and recapture intentionally"
                )
        if args.capture_runtime and not errors:
            RUNTIME.write_text(json.dumps({
                "provenance": "Local FastMCP registration; no COM initialized or mail accessed",
                "toolCount": len(runtime),
                "tools": runtime,
            }, indent=2, sort_keys=True, ensure_ascii=False) + "\n", encoding="utf-8")
            print(f"Wrote {RUNTIME.relative_to(ROOT)} (no mailbox content)")
    if errors:
        for error in errors:
            print("FAIL:", error, file=sys.stderr)
        return 1
    print(f"PASS: {EXPECTED} Python declarations" +
          (" (source only)" if args.static else " (FastMCP metadata; no Outlook testing)"))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
