from __future__ import annotations

import os
import urllib.request


def main() -> int:
    for key in list(os.environ):
        if key.lower() in {
            "http_proxy", "https_proxy", "all_proxy", "ftp_proxy",
            "pip_proxy", "pip_index_url", "pip_extra_index_url",
            "pip_trusted_host",
        }:
            os.environ.pop(key, None)

    os.environ["NO_PROXY"] = "*"
    os.environ["no_proxy"] = "*"
    os.environ["PIP_CONFIG_FILE"] = os.devnull
    os.environ["PIP_DISABLE_PIP_VERSION_CHECK"] = "1"

    urllib.request.getproxies = lambda: {}
    urllib.request.getproxies_environment = lambda: {}
    urllib.request.proxy_bypass = lambda host: True

    from pip._internal.cli.main import main as pip_main

    return int(
        pip_main(
            [
                "install",
                "--isolated",
                "--no-cache-dir",
                "--disable-pip-version-check",
                "--index-url",
                "https://pypi.org/simple",
                "-r",
                "requirements.txt",
            ]
        )
    )


if __name__ == "__main__":
    raise SystemExit(main())
