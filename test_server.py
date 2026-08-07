from __future__ import annotations

import sys


def test_imports() -> None:
    print("Testing imports...")

    import catia_mcp
    from catia_mcp.connection import CATIAConnection
    from catia_mcp.registry import ToolContext
    from catia_mcp.server import CATIAMCPServer

    print("  Imports OK.")


def test_server_creation() -> int:
    print("Testing server creation...")

    from catia_mcp.server import CATIAMCPServer

    server = CATIAMCPServer()
    loaded = server._registered_tools

    assert loaded, "No MCP tools were loaded."

    duplicates = sorted({name for name in loaded if loaded.count(name) > 1})
    assert not duplicates, f"Duplicate tools found: {duplicates}"

    for name in loaded:
        assert name.startswith("catia_"), f"Tool name must start with catia_: {name}"

    print(f"  Loaded {len(loaded)} tools.")
    for name in loaded:
        print(f"    - {name}")

    return len(loaded)


def main() -> int:
    print("=" * 70)
    print("CATIA V5 MCP Server - Import and Registration Test")
    print("=" * 70)

    try:
        test_imports()
        count = test_server_creation()

        print("=" * 70)
        print(f"ALL TESTS PASSED - {count} tools registered")
        print("=" * 70)
        return 0

    except Exception as exc:
        print(f"TEST FAILED: {exc}")
        import traceback

        traceback.print_exc()
        return 1


if __name__ == "__main__":
    sys.exit(main())
