from __future__ import annotations

import os
import shutil
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
VENV = ROOT / ".venv"
LOG = ROOT / "install_log.txt"


def write(line: str) -> None:
    print(line)
    with LOG.open("a", encoding="utf-8") as fh:
        fh.write(line + "\n")


def run(command: list[str], env: dict[str, str] | None = None) -> None:
    write("RUN: " + " ".join(command))
    result = subprocess.run(
        command,
        cwd=ROOT,
        env=env,
        text=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        errors="replace",
    )
    if result.stdout:
        print(result.stdout, end="")
        with LOG.open("a", encoding="utf-8") as fh:
            fh.write(result.stdout)
    if result.returncode != 0:
        raise RuntimeError(f"Command failed with code {result.returncode}")


def clean_proxy_env() -> dict[str, str]:
    env = os.environ.copy()
    for key in list(env):
        if key.lower() in {
            "http_proxy", "https_proxy", "all_proxy", "ftp_proxy",
            "pip_proxy", "pip_index_url", "pip_extra_index_url",
            "pip_trusted_host",
        }:
            env.pop(key, None)
    env["NO_PROXY"] = "*"
    env["no_proxy"] = "*"
    env["PIP_CONFIG_FILE"] = os.devnull
    env["PIP_DISABLE_PIP_VERSION_CHECK"] = "1"
    return env


def main() -> int:
    LOG.write_text("MAGNIT RESET INSTALL LOG\n", encoding="utf-8")
    write(f"Python: {sys.version}")
    write(f"Folder: {ROOT}")

    if VENV.exists():
        write("Removing old .venv...")
        shutil.rmtree(VENV, ignore_errors=True)
        if VENV.exists():
            raise RuntimeError("Could not remove .venv. Close Python and try again.")

    write("Creating .venv...")
    run([sys.executable, "-m", "venv", str(VENV)])

    vpy = VENV / "Scripts" / "python.exe"
    if not vpy.exists():
        raise RuntimeError("Virtual environment Python was not created.")

    write("Checking pip...")
    run([str(vpy), "-m", "ensurepip", "--upgrade"])

    write("Installing packages with direct connection...")
    env = clean_proxy_env()
    run([str(vpy), str(ROOT / "pip_direct.py")], env=env)

    write("Testing imports...")
    run([
        str(vpy),
        "-c",
        "import playwright, openpyxl; print('IMPORTS OK')",
    ], env=env)

    write("INSTALL COMPLETE")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        write("ERROR: " + repr(exc))
        raise SystemExit(1)
