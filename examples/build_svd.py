"""Example wrapper: gather release data, then render the SVD.

A real wrapper would pull the version from git tags, hash the actual build
outputs, and query Jira for the issues fixed in this release. Here the
release directory is optional; without one, sample data is used.

    python examples/build_svd.py [DIST_DIR] [OUTPUT.docx]
"""

import hashlib
import sys
from pathlib import Path

from rst_to_docx import render

HERE = Path(__file__).parent


def file_records(dist):
    return [
        {
            "name": path.name,
            "version": "1.4.2",
            "size": f"{path.stat().st_size:,}",
            "sha256": hashlib.sha256(path.read_bytes()).hexdigest(),
        }
        for path in sorted(dist.iterdir())
        if path.is_file()
    ]


SAMPLE_FILES = [
    {"name": "widget-1.4.2-setup.exe", "version": "1.4.2", "size": "48,211,456",
     "sha256": "9f2c4a1be7d03f6a5c8e1b2d4f6a8c0e2b4d6f8a0c2e4b6d8f0a2c4e6b8d0f2a"},
    {"name": "widget-fw-3.1.0.bin", "version": "3.1.0", "size": "1,048,576",
     "sha256": "1b3d5f7a9c1e3b5d7f9a1c3e5b7d9f1a3c5e7b9d1f3a5c7e9b1d3f5a7c9e1b3d"},
]

CHANGELOG = [
    {"key": "SVD-101", "type": "Bug", "summary": "Installer fails when the path contains spaces."},
    {"key": "SVD-117", "type": "Feature", "summary": "Add silent-install option /quiet."},
    {"key": ["SVD-120", "SVD-121"], "type": "Bug", "summary": "Firmware update retries on timeout."},
]


def main():
    dist = Path(sys.argv[1]) if len(sys.argv) > 1 else None
    output = Path(sys.argv[2]) if len(sys.argv) > 2 else HERE / "svd.docx"
    template = HERE / "template.docx"
    if not template.exists():
        from make_template import build
        build(template)

    variables = {
        "project": "Widget Controller",
        "release": {"version": "1.4.2", "previous": "1.4.1", "date": "2026-10-13"},
        "doc": {"number": "SVD-WC-0042", "author": "Release Engineering",
                "srs": "SRS-WC-0007", "sdd": "SDD-WC-0011"},
        "build": {"compiler": "GCC 13.2", "commit": "57eb69e", "id": "1832"},
        "files": file_records(dist) if dist else SAMPLE_FILES,
        "changelog": CHANGELOG,
    }
    render(HERE / "svd.rst", template, output,
           variables=variables, jira_url="https://jira.example.com/browse/")
    print(output)


if __name__ == "__main__":
    main()
