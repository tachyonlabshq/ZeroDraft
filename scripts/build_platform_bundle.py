#!/usr/bin/env python3
from __future__ import annotations

import argparse
import hashlib
import json
import shutil
import tempfile
import zipfile
from pathlib import Path


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def render_template(path: Path, replacements: dict[str, str]) -> str:
    content = path.read_text(encoding="utf-8")
    for key, value in replacements.items():
        content = content.replace(key, value)
    return content


def build_bundle(args: argparse.Namespace) -> None:
    binary_path = Path(args.binary_path).resolve()
    if not binary_path.exists():
        raise FileNotFoundError(f"binary not found: {binary_path}")

    readme_path = Path(args.readme_path).resolve()
    skill_path = Path(args.skill_path).resolve()
    mcp_template_path = Path(args.mcp_template_path).resolve()
    output_root = Path(args.output_root).resolve()
    output_root.mkdir(parents=True, exist_ok=True)

    bundle_basename = f"{args.project_name}-{args.platform}-{args.version}"
    bundle_dir_name = args.project_name
    bundled_binary_name = binary_path.name

    with tempfile.TemporaryDirectory(prefix=f"{bundle_basename}-") as tmpdir:
        staging_root = Path(tmpdir) / bundle_dir_name
        staging_bin = staging_root / "bin"
        staging_bin.mkdir(parents=True, exist_ok=True)

        shutil.copy2(readme_path, staging_root / "README.md")
        shutil.copy2(skill_path, staging_root / "SKILL.md")
        shutil.copy2(binary_path, staging_bin / bundled_binary_name)

        mcp_json = render_template(
            mcp_template_path,
            {
                "__SKILL_KEY__": args.skill_key,
                "__BINARY_PATH__": f"./bin/{bundled_binary_name}",
            },
        )
        (staging_root / "mcp.json").write_text(mcp_json + "\n", encoding="utf-8")

        zip_path = output_root / f"{bundle_basename}.zip"
        with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            for file_path in sorted(staging_root.rglob("*")):
                archive.write(file_path, file_path.relative_to(Path(tmpdir)))

    zip_digest = sha256_bytes(zip_path.read_bytes())
    manifest = {
        "project_name": args.project_name,
        "skill_key": args.skill_key,
        "platform": args.platform,
        "version": args.version,
        "archive_name": zip_path.name,
        "bundle_folder": bundle_dir_name,
        "binary_name": bundled_binary_name,
        "binary_relative_path": f"bin/{bundled_binary_name}",
        "mcp_command": [f"./bin/{bundled_binary_name}", "mcp-stdio"],
        "files": [
            f"{bundle_dir_name}/README.md",
            f"{bundle_dir_name}/SKILL.md",
            f"{bundle_dir_name}/mcp.json",
            f"{bundle_dir_name}/bin/{bundled_binary_name}",
        ],
        "sha256": zip_digest,
    }

    manifest_path = output_root / f"{bundle_basename}.manifest.json"
    manifest_path.write_text(json.dumps(manifest, indent=2) + "\n", encoding="utf-8")

    checksum_path = output_root / f"{bundle_basename}.sha256.txt"
    checksum_path.write_text(f"{zip_digest}  {zip_path.name}\n", encoding="utf-8")

    print(
        json.dumps(
            {
                "zip_path": str(zip_path),
                "manifest_path": str(manifest_path),
                "checksum_path": str(checksum_path),
            }
        )
    )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Build a self-contained, install-ready platform bundle for a Zero-family skill."
    )
    parser.add_argument("--platform", required=True)
    parser.add_argument("--binary-path", required=True)
    parser.add_argument("--output-root", required=True)
    parser.add_argument("--project-name", default="ZeroDraft")
    parser.add_argument("--skill-key", default="zerodraft")
    parser.add_argument("--version", default="dev")
    parser.add_argument("--readme-path", default="README.md")
    parser.add_argument("--skill-path", default="SKILL.md")
    parser.add_argument(
        "--mcp-template-path",
        default="distribution/templates/platform-package-mcp.template.json",
    )
    return parser.parse_args()


if __name__ == "__main__":
    build_bundle(parse_args())
