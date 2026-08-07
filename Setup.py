import asyncio
import importlib
import inspect
import json
import os
import shutil
import subprocess
import sys
from pathlib import Path
from typing import Optional


APP_NAME = "CATIA MCP Setup for Codex"
SERVER_NAME = "catia-mcp"
SCRIPT_DIR = Path(__file__).resolve().parent
REQUIRED_PACKAGES = {
    "pythoncom": "pywin32",
    "win32com.client": "pywin32",
}


def print_header(title: str) -> None:
    print("\n" + "=" * 68)
    print(title)
    print("=" * 68)


def run_command(
    command: list[str],
    *,
    cwd: Optional[Path] = None,
    timeout: int = 120,
    capture_output: bool = True,
) -> subprocess.CompletedProcess:
    return subprocess.run(
        command,
        cwd=str(cwd) if cwd else None,
        text=True,
        encoding="utf-8",
        errors="replace",
        capture_output=capture_output,
        timeout=timeout,
        check=False,
    )


def ensure_pip() -> bool:
    result = run_command([sys.executable, "-m", "pip", "--version"], timeout=30)
    if result.returncode == 0:
        return True

    print("[INFO] pip is not available. Attempting to enable it with ensurepip...")
    result = run_command([sys.executable, "-m", "ensurepip", "--upgrade"], timeout=120)
    if result.returncode != 0:
        print("[ERROR] Unable to enable pip.")
        if result.stderr:
            print(result.stderr.strip())
        return False
    return True


def module_is_available(module_name: str) -> bool:
    try:
        importlib.import_module(module_name)
        return True
    except ImportError:
        return False


def install_packages() -> bool:
    print_header("Checking Python dependencies")

    if not ensure_pip():
        return False

    missing_packages = sorted(
        {
            package_name
            for module_name, package_name in REQUIRED_PACKAGES.items()
            if not module_is_available(module_name)
        }
    )

    if not missing_packages:
        print("[OK] Required Python packages are already installed.")
    else:
        print(f"[INFO] Missing packages: {', '.join(missing_packages)}")
        command = [
            sys.executable,
            "-m",
            "pip",
            "install",
            "--disable-pip-version-check",
            *missing_packages,
        ]
        result = run_command(command, timeout=300)

        if result.returncode != 0 and sys.prefix == sys.base_prefix:
            print("[INFO] System-wide installation failed. Retrying for the current user...")
            command.insert(4, "--user")
            result = run_command(command, timeout=300)

        if result.returncode != 0:
            print("[ERROR] Failed to install required Python packages.")
            if result.stdout:
                print(result.stdout.strip())
            if result.stderr:
                print(result.stderr.strip())
            return False

        importlib.invalidate_caches()
        still_missing = [
            module_name
            for module_name in REQUIRED_PACKAGES
            if not module_is_available(module_name)
        ]
        if still_missing:
            print(
                "[ERROR] Package installation completed, but these modules still "
                f"cannot be imported: {', '.join(still_missing)}"
            )
            print("[INFO] Restart Python or verify that pip is installing into the same interpreter.")
            return False

        print("[OK] Required Python packages were installed successfully.")

    requirements_file = SCRIPT_DIR / "requirements.txt"
    if requirements_file.is_file():
        print(f"[INFO] Installing project dependencies from {requirements_file.name}...")
        result = run_command(
            [
                sys.executable,
                "-m",
                "pip",
                "install",
                "--disable-pip-version-check",
                "-r",
                str(requirements_file),
            ],
            cwd=SCRIPT_DIR,
            timeout=600,
        )
        if result.returncode != 0:
            print("[ERROR] Failed to install dependencies from requirements.txt.")
            if result.stderr:
                print(result.stderr.strip())
            return False
        print("[OK] Project dependencies are installed.")

    return True


def check_environment() -> bool:
    print_header("Checking Windows and CATIA environment")

    if sys.platform != "win32":
        print("[ERROR] CATIA V5 COM automation requires Windows.")
        return False

    if sys.version_info < (3, 9):
        print(
            f"[ERROR] Python {sys.version_info.major}.{sys.version_info.minor} is too old. "
            "Use Python 3.9 or newer."
        )
        return False

    catia_package = SCRIPT_DIR / "catia_mcp"
    if not catia_package.is_dir():
        print(f"[ERROR] Local package directory was not found: {catia_package}")
        print("[INFO] Place Setup.py next to the 'catia_mcp' directory.")
        return False

    if str(SCRIPT_DIR) not in sys.path:
        sys.path.insert(0, str(SCRIPT_DIR))

    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        print(f"[ERROR] pywin32 is not usable: {exc}")
        return False

    try:
        _ = win32com.client.Dispatch("CATIA.Application")
        print("[OK] CATIA.Application is registered as a COM server.")
    except pythoncom.com_error as exc:
        print(f"[ERROR] CATIA.Application is not registered correctly: {exc}")
        print("[INFO] Verify that CATIA V5 is installed and its COM registration is intact.")
        return False
    except Exception as exc:
        print(f"[ERROR] Unexpected error checking CATIA COM registration: {exc}")
        return False

    initialized = False
    try:
        pythoncom.CoInitialize()
        initialized = True
        try:
            win32com.client.GetActiveObject("CATIA.Application")
            print("[OK] A running CATIA instance was detected.")
        except Exception:
            print("[INFO] CATIA is installed, but no running CATIA instance was detected.")
            print("[INFO] This is not an error. The MCP server can connect when CATIA is started.")
    finally:
        if initialized:
            pythoncom.CoUninitialize()

    try:
        module = importlib.import_module("catia_mcp.server")
        server_class = getattr(module, "CATIAMCPServer")
        if not inspect.isclass(server_class):
            raise TypeError("CATIAMCPServer is not a class")
        print("[OK] catia_mcp.server.CATIAMCPServer can be imported.")
    except Exception as exc:
        print(f"[ERROR] Unable to import the local CATIA MCP server: {exc}")
        print("[INFO] Check the catia_mcp package and its Python dependencies.")
        return False

    return True


async def _await_if_needed(value) -> None:
    if inspect.isawaitable(value):
        await value


def test_mcp_server() -> bool:
    print_header("Testing CATIA MCP initialization")

    try:
        from catia_mcp.server import CATIAMCPServer

        server = CATIAMCPServer()
        setup_method = getattr(server, "setup", None)
        if callable(setup_method):
            result = setup_method()
            if inspect.isawaitable(result):
                asyncio.run(_await_if_needed(result))

        run_method = getattr(server, "run", None)
        if not callable(run_method):
            print("[ERROR] CATIAMCPServer does not expose a callable run() method.")
            return False

        print("[OK] CATIA MCP server initialized successfully.")
        return True
    except Exception as exc:
        print(f"[ERROR] CATIA MCP initialization failed: {exc}")
        return False


def find_codex() -> Optional[str]:
    candidates = ["codex", "codex.exe", "codex.cmd"]
    for candidate in candidates:
        found = shutil.which(candidate)
        if found:
            return found
    return None


def install_codex_cli() -> Optional[str]:
    print_header("Checking Codex CLI")

    codex_path = find_codex()
    if codex_path:
        result = run_command([codex_path, "--version"], timeout=30)
        version = (result.stdout or result.stderr).strip()
        print(f"[OK] Codex CLI found: {version or codex_path}")
        return codex_path

    print("[INFO] Codex CLI was not found in PATH.")

    npm_path = shutil.which("npm") or shutil.which("npm.cmd")
    if npm_path:
        print("[INFO] npm was found. Installing the official @openai/codex package...")
        result = run_command([npm_path, "install", "-g", "@openai/codex"], timeout=600)
        if result.returncode != 0:
            print("[ERROR] npm could not install Codex CLI.")
            if result.stderr:
                print(result.stderr.strip())
            return None

        codex_path = find_codex()
        if codex_path:
            result = run_command([codex_path, "--version"], timeout=30)
            version = (result.stdout or result.stderr).strip()
            print(f"[OK] Codex CLI installed successfully: {version or codex_path}")
            return codex_path

    print("[ERROR] Codex CLI is required but could not be installed automatically.")
    print("[INFO] Install Codex CLI with the official Windows installer, then rerun this script.")
    print(
        "[INFO] Official command: powershell -ExecutionPolicy ByPass -c "
        '"irm https://chatgpt.com/codex/install.ps1 | iex"'
    )
    return None


def create_launcher() -> Path:
    launcher_path = SCRIPT_DIR / ".catia_mcp_codex_launcher.py"
    launcher_code = f'''# Auto-generated by Setup.py.\nimport os\nimport runpy\nimport sys\nfrom pathlib import Path\n\nROOT = Path({str(SCRIPT_DIR)!r})\nos.chdir(ROOT)\nif str(ROOT) not in sys.path:\n    sys.path.insert(0, str(ROOT))\nrunpy.run_module("catia_mcp", run_name="__main__")\n'''
    launcher_path.write_text(launcher_code, encoding="utf-8")
    return launcher_path


def configure_codex_mcp(codex_path: str) -> bool:
    print_header("Configuring CATIA MCP in Codex")

    launcher_path = create_launcher()
    print(f"[INFO] MCP launcher: {launcher_path}")

    existing = run_command([codex_path, "mcp", "get", SERVER_NAME, "--json"], timeout=30)
    if existing.returncode == 0:
        print(f"[INFO] Existing '{SERVER_NAME}' configuration found. Replacing it safely...")
        remove_result = run_command([codex_path, "mcp", "remove", SERVER_NAME], timeout=30)
        if remove_result.returncode != 0:
            print("[ERROR] Could not remove the existing CATIA MCP configuration.")
            if remove_result.stderr:
                print(remove_result.stderr.strip())
            return False

    add_command = [
        codex_path,
        "mcp",
        "add",
        SERVER_NAME,
        "--env",
        "CATIA_MCP_LOG_LEVEL=INFO",
        "--",
        sys.executable,
        "-B",
        str(launcher_path),
    ]
    add_result = run_command(add_command, timeout=60)
    if add_result.returncode != 0:
        print("[ERROR] Codex failed to register the CATIA MCP server.")
        if add_result.stdout:
            print(add_result.stdout.strip())
        if add_result.stderr:
            print(add_result.stderr.strip())
        return False

    verify = run_command([codex_path, "mcp", "get", SERVER_NAME, "--json"], timeout=30)
    if verify.returncode != 0:
        print("[ERROR] Codex reported success, but the MCP configuration cannot be read back.")
        if verify.stderr:
            print(verify.stderr.strip())
        return False

    try:
        config = json.loads(verify.stdout)
        transport = config.get("transport", config)
        command = transport.get("command")
        args = transport.get("args", [])
        enabled = config.get("enabled", True)

        expected_launcher = str(launcher_path)
        if os.path.normcase(os.path.abspath(command or "")) != os.path.normcase(
            os.path.abspath(sys.executable)
        ):
            raise ValueError(f"Unexpected command in Codex config: {command}")
        if expected_launcher not in args:
            raise ValueError("The generated CATIA MCP launcher is missing from Codex arguments.")
        if not enabled:
            raise ValueError("The CATIA MCP server is registered but disabled.")
    except (json.JSONDecodeError, TypeError, ValueError) as exc:
        print(f"[ERROR] Codex MCP verification failed: {exc}")
        print("[INFO] Raw Codex output:")
        print(verify.stdout.strip())
        return False

    print("[OK] CATIA MCP is registered and enabled in Codex.")
    print("[INFO] Codex CLI and the Codex IDE extension share the same MCP configuration.")
    print("[INFO] Restart any already-open Codex session before using the new MCP server.")
    return True


def show_codex_status(codex_path: str) -> None:
    result = run_command([codex_path, "mcp", "list"], timeout=30)
    if result.returncode == 0 and result.stdout.strip():
        print("\nCurrent Codex MCP configuration:")
        print(result.stdout.strip())


def start_mcp_server() -> bool:
    print_header("Starting CATIA MCP server directly")
    print("[INFO] Press Ctrl+C to stop the server.")

    if str(SCRIPT_DIR) not in sys.path:
        sys.path.insert(0, str(SCRIPT_DIR))
    os.chdir(SCRIPT_DIR)

    try:
        from catia_mcp.server import CATIAMCPServer

        server = CATIAMCPServer()
        server.run()
        return True
    except KeyboardInterrupt:
        print("\n[OK] CATIA MCP server stopped.")
        return True
    except Exception as exc:
        print(f"[ERROR] CATIA MCP server failed: {exc}")
        return False


def pause_before_exit() -> None:
    if sys.stdin.isatty():
        try:
            input("\nPress Enter to exit...")
        except EOFError:
            pass


def main() -> int:
    print_header(APP_NAME)
    print(f"Python executable: {sys.executable}")
    print(f"Setup directory:   {SCRIPT_DIR}")

    if not install_packages():
        return 1

    if not check_environment():
        return 1

    if not test_mcp_server():
        return 1

    codex_path = install_codex_cli()
    if not codex_path:
        return 1

    if not configure_codex_mcp(codex_path):
        return 1

    show_codex_status(codex_path)

    print_header("Setup complete")
    print("[OK] CATIA MCP has been configured for Codex successfully.")
    print("[INFO] Start CATIA, open a new Codex session, and check the MCP tools there.")
    return 0


if __name__ == "__main__":
    exit_code = main()
    pause_before_exit()
    raise SystemExit(exit_code)
