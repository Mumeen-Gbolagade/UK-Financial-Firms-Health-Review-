"""
Project settings shared by the build scripts.

REPO_URL is the GitHub repository this project is published in. Links in the
Excel workbook and the report point here, so they keep working after someone
downloads a single file. Set the REPO_URL environment variable to override it
(for example, after forking the repository).
"""

import os

REPO_URL = os.environ.get(
    "REPO_URL", "https://github.com/Mumeen-Gbolagade/UK-Financial-Firms-Health-Review-"
).rstrip("/")
BRANCH = os.environ.get("REPO_BRANCH", "main")


def github_link(path: str) -> str:
    """GitHub page for a file or folder in this repository, e.g. 'report/x.pdf'."""
    path = path.strip("/")
    kind = "blob" if "." in path.rsplit("/", 1)[-1] else "tree"
    return f"{REPO_URL}/{kind}/{BRANCH}/{path}"
