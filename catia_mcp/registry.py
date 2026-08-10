from __future__ import annotations

import importlib
import logging
import pkgutil
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from catia_mcp.connection import CATIAConnection

logger = logging.getLogger("catia_mcp.registry")


@dataclass
class ToolContext:
    conn: CATIAConnection


def get_tool_modules() -> list[str]:
    """Discover importable tool modules in ``catia_mcp.tools``."""
    tools_dir = Path(__file__).parent / "tools"
    tool_modules: list[str] = []

    for _, name, is_pkg in pkgutil.iter_modules([str(tools_dir)]):
        if not is_pkg and not name.startswith("_"):
            tool_modules.append(name)

    # A stable order makes startup logs and registration behavior reproducible.
    return sorted(tool_modules)


TOOL_MODULES = get_tool_modules()


def load_tool_modules(
    mcp_instance: Any,
    ctx_instance: ToolContext,
) -> list[str]:
    """Import and register tool modules without polluting MCP stdout.

    An individual module import/registration failure is logged to ``stderr``
    and does not prevent the remaining modules or the MCP stdio server from
    starting.
    """
    loaded_tools: list[str] = []

    for module_name in TOOL_MODULES:
        qualified_name = f".tools.{module_name}"

        try:
            module = importlib.import_module(
                qualified_name,
                package=__package__,
            )
        except Exception:
            logger.exception(
                "Failed to import tool module '%s'; skipping it.",
                module_name,
            )
            continue

        register_tools = getattr(module, "register_tools", None)
        if not callable(register_tools):
            logger.warning(
                "Tool module '%s' has no callable register_tools(); skipping it.",
                module_name,
            )
            continue

        try:
            result = register_tools(mcp_instance, ctx_instance)
        except Exception:
            logger.exception(
                "Failed to register tools from module '%s'; skipping it.",
                module_name,
            )
            continue

        if result is None:
            logger.info(
                "Tool module '%s' completed registration without returning names.",
                module_name,
            )
            continue

        if isinstance(result, list):
            module_tools = [str(name) for name in result]
        else:
            # Preserve compatibility with older modules that returned one name.
            module_tools = [str(result)]
            logger.warning(
                "Tool module '%s' returned %s instead of list[str].",
                module_name,
                type(result).__name__,
            )

        loaded_tools.extend(module_tools)
        logger.info(
            "Registered %d tool(s) from module '%s'.",
            len(module_tools),
            module_name,
        )

    logger.info("Tool discovery completed: %d tool(s) registered.", len(loaded_tools))
    return loaded_tools
