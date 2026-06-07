#!/usr/bin/env python3
"""
run_streamlit_uv.py

Place this file in the same folder as your Streamlit app, then run:

    python run_streamlit_uv.py

What it does:
1. Installs uv if uv is not available.
2. Initializes a uv project if pyproject.toml does not exist.
3. Creates .venv if it does not exist.
4. Uses uv to install required packages into .venv if missing.
5. Runs playwright install after Playwright is available.
6. Runs the Streamlit app and opens it in the browser.
"""

from __future__ import annotations

import argparse
import os
import shutil
import site
import subprocess
import sys
from pathlib import Path


# Packages needed by your Streamlit app.
# Add/remove packages here if your project grows.
REQUIRED_PACKAGES = [
    "streamlit",
    "pandas",
    "openpyxl",
    "playwright",  # used by the automation file in your app flow
]

# Package name -> import name. Most packages use the same name.
IMPORT_NAMES = {
    "streamlit": "streamlit",
    "pandas": "pandas",
    "openpyxl": "openpyxl",
    "playwright": "playwright",
}

APP_CANDIDATES = [
    "app.py",
    "app_blank_email_allowed.py",
    "app_table_reflect_fixed.py",
    "app_color_updated.py",
    "Pasted code.py",
]


class SetupError(RuntimeError):
    """Raised when setup cannot continue."""


def step(message: str) -> None:
    print(f"\n==> {message}")


def run_command(
    command: list[str],
    cwd: Path,
    env: dict[str, str] | None = None,
    check: bool = True,
) -> subprocess.CompletedProcess[str]:
    print("$ " + " ".join(str(part) for part in command))
    return subprocess.run(
        command,
        cwd=str(cwd),
        env=env,
        text=True,
        check=check,
    )


def user_scripts_dir() -> Path:
    """Return the user-level Python scripts directory for PATH refresh after pip install --user."""
    base = Path(site.USER_BASE)
    if os.name == "nt":
        return base / "Scripts"
    return base / "bin"


def get_uv_command() -> list[str] | None:
    """Return a uv command if uv is available, otherwise None."""
    uv_exe = shutil.which("uv")
    if uv_exe:
        return [uv_exe]

    # Fallback for uv installed as a Python module.
    try:
        subprocess.run(
            [sys.executable, "-m", "uv", "--version"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            check=True,
        )
        return [sys.executable, "-m", "uv"]
    except Exception:
        return None


def ensure_uv(project_dir: Path) -> list[str]:
    """Install uv if needed and return the command used to invoke it."""
    step("Checking uv")
    uv_cmd = get_uv_command()
    if uv_cmd:
        run_command(uv_cmd + ["--version"], cwd=project_dir)
        return uv_cmd

    step("uv not found. Installing uv with pip")
    run_command([sys.executable, "-m", "pip", "install", "--user", "uv"], cwd=project_dir)

    # Make the newly installed uv executable visible to this running script.
    scripts_dir = user_scripts_dir()
    os.environ["PATH"] = str(scripts_dir) + os.pathsep + os.environ.get("PATH", "")

    uv_cmd = get_uv_command()
    if not uv_cmd:
        raise SetupError(
            "uv was installed, but the uv command is still not available. "
            f"Try adding this directory to PATH: {scripts_dir}"
        )

    run_command(uv_cmd + ["--version"], cwd=project_dir)
    return uv_cmd


def ensure_uv_project(uv_cmd: list[str], project_dir: Path) -> None:
    """Create pyproject.toml with uv init if the folder is not already a uv project."""
    pyproject = project_dir / "pyproject.toml"
    if pyproject.exists():
        step("uv project already initialized")
        print(f"Found: {pyproject}")
        return

    step("Initializing uv project")
    run_command(
        uv_cmd
        + [
            "init",
            "--bare",
            "--app",
            "--name",
            "v2vedtech-admin",
            "--no-workspace",
        ],
        cwd=project_dir,
    )


def venv_python_path(venv_dir: Path) -> Path:
    if os.name == "nt":
        return venv_dir / "Scripts" / "python.exe"
    return venv_dir / "bin" / "python"


def venv_bin_dir(venv_dir: Path) -> Path:
    if os.name == "nt":
        return venv_dir / "Scripts"
    return venv_dir / "bin"


def ensure_venv(uv_cmd: list[str], project_dir: Path, venv_dir: Path) -> Path:
    """Create .venv if missing and return its Python executable."""
    py = venv_python_path(venv_dir)
    if py.exists():
        step("Virtual environment already exists")
        print(f"Found: {venv_dir}")
        return py

    step("Creating virtual environment with uv")
    run_command(uv_cmd + ["venv", str(venv_dir)], cwd=project_dir)

    if not py.exists():
        raise SetupError(f"Virtual environment was created, but Python was not found at: {py}")
    return py


def activated_env(venv_dir: Path) -> dict[str, str]:
    """
    Build an activated-venv environment for child processes.

    A Python script cannot permanently activate a venv in the parent terminal,
    so this sets VIRTUAL_ENV and PATH for commands launched by this script.
    """
    env = os.environ.copy()
    env["VIRTUAL_ENV"] = str(venv_dir)
    env["PATH"] = str(venv_bin_dir(venv_dir)) + os.pathsep + env.get("PATH", "")
    env.pop("PYTHONHOME", None)
    return env


def package_is_installed(venv_py: Path, import_name: str, project_dir: Path, env: dict[str, str]) -> bool:
    result = subprocess.run(
        [str(venv_py), "-c", f"import {import_name}"],
        cwd=str(project_dir),
        env=env,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        text=True,
    )
    return result.returncode == 0


def ensure_packages(
    uv_cmd: list[str],
    project_dir: Path,
    venv_py: Path,
    env: dict[str, str],
    packages: list[str],
) -> None:
    """Install missing Python packages into .venv using uv."""
    step("Checking Python packages")
    missing: list[str] = []

    for package in packages:
        import_name = IMPORT_NAMES.get(package, package.replace("-", "_"))
        if package_is_installed(venv_py, import_name, project_dir, env):
            print(f"OK: {package}")
        else:
            print(f"Missing: {package}")
            missing.append(package)

    if not missing:
        print("All required packages are already installed.")
        return

    step("Installing missing packages with uv")
    run_command(
        uv_cmd + ["pip", "install", "--python", str(venv_py)] + missing,
        cwd=project_dir,
        env=env,
    )


def ensure_playwright_browsers(project_dir: Path, venv_py: Path, env: dict[str, str]) -> None:
    """
    Run Playwright's browser installer after the playwright package is available.

    The command is safe to run again. If browsers are already installed,
    Playwright reuses the existing browser files.
    """
    if not package_is_installed(venv_py, "playwright", project_dir, env):
        print("Playwright package is not installed, so skipping playwright install.")
        return

    step("Installing Playwright browsers")
    run_command([str(venv_py), "-m", "playwright", "install"], cwd=project_dir, env=env)


def resolve_app_file(project_dir: Path, app_arg: str | None, launcher_path: Path) -> Path:
    if app_arg:
        app_path = Path(app_arg)
        if not app_path.is_absolute():
            app_path = project_dir / app_path
        if app_path.exists():
            return app_path
        raise SetupError(f"Streamlit app file not found: {app_path}")

    for name in APP_CANDIDATES:
        candidate = project_dir / name
        if candidate.exists() and candidate.resolve() != launcher_path.resolve():
            return candidate

    py_files = [
        p
        for p in project_dir.glob("*.py")
        if p.resolve() != launcher_path.resolve() and p.name != "automations.py"
    ]
    if len(py_files) == 1:
        return py_files[0]

    raise SetupError(
        "Could not automatically find your Streamlit app. "
        "Run this script with --app app.py, or rename your Streamlit file to app.py."
    )


def run_streamlit(project_dir: Path, venv_py: Path, app_file: Path, env: dict[str, str]) -> int:
    step("Starting Streamlit")
    print(f"App file: {app_file}")
    print("Streamlit should open in your browser. If it does not, copy the local URL from the terminal.")

    command = [
        str(venv_py),
        "-m",
        "streamlit",
        "run",
        str(app_file),
        "--server.headless=false",
    ]
    return subprocess.call(command, cwd=str(project_dir), env=env)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Set up uv/.venv and run a Streamlit app.")
    parser.add_argument(
        "--app",
        default=None,
        help="Streamlit app file to run. Example: --app app.py",
    )
    parser.add_argument(
        "--project-dir",
        default=None,
        help="Project directory. Defaults to the folder containing this script.",
    )
    parser.add_argument(
        "--skip-playwright-install",
        action="store_true",
        help="Skip running: python -m playwright install",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    launcher_path = Path(__file__).resolve()
    project_dir = Path(args.project_dir).resolve() if args.project_dir else launcher_path.parent
    venv_dir = project_dir / ".venv"

    try:
        uv_cmd = ensure_uv(project_dir)
        ensure_uv_project(uv_cmd, project_dir)
        venv_py = ensure_venv(uv_cmd, project_dir, venv_dir)
        env = activated_env(venv_dir)
        ensure_packages(uv_cmd, project_dir, venv_py, env, REQUIRED_PACKAGES)

        if not args.skip_playwright_install:
            ensure_playwright_browsers(project_dir, venv_py, env)

        app_file = resolve_app_file(project_dir, args.app, launcher_path)
        return run_streamlit(project_dir, venv_py, app_file, env)

    except subprocess.CalledProcessError as exc:
        print(f"\nCommand failed with exit code {exc.returncode}.")
        return exc.returncode
    except SetupError as exc:
        print(f"\nSetup error: {exc}")
        return 1
    except KeyboardInterrupt:
        print("\nStopped by user.")
        return 130


if __name__ == "__main__":
    raise SystemExit(main())
