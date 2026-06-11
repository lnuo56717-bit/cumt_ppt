from __future__ import annotations

import argparse
import hashlib
import json
import sys
import traceback
from pathlib import Path
from typing import Any, Callable

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
sys.path.insert(0, str(SRC))

from cumt_ppt_mcp.ppt_utils import check_ppt_quality, inspect_ppt, inspect_slide  # noqa: E402


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def run_tool(name: str, fn: Callable[..., Any], *args: Any, **kwargs: Any) -> dict[str, Any]:
    try:
        value = fn(*args, **kwargs)
        return {"tool": name, "success": bool(value.get("ok", True) if isinstance(value, dict) else True), "result": value}
    except Exception as exc:
        return {"tool": name, "success": False, "error": f"{type(exc).__name__}: {exc}", "traceback": traceback.format_exc()}


def main() -> int:
    parser = argparse.ArgumentParser(description="Run optional local regression checks on a user-supplied PPTX.")
    parser.add_argument("pptx_path", help="Path to a local PPTX. Do not commit this file or generated outputs.")
    parser.add_argument("--output-dir", default=str(ROOT / "test_outputs"), help="Directory for local-only reports.")
    parser.add_argument("--expected-title", default=None)
    parser.add_argument("--required-number", action="append", default=[])
    parser.add_argument("--inspect-slide", action="append", type=int, default=[])
    args = parser.parse_args()

    pptx = Path(args.pptx_path).expanduser().resolve()
    out = Path(args.output_dir).expanduser().resolve()
    out.mkdir(parents=True, exist_ok=True)
    before_hash = sha256(pptx)
    results: dict[str, Any] = {"pptx_path": str(pptx), "before_hash": before_hash, "tools": {}, "focused_slides": {}}

    results["tools"]["inspect_ppt"] = run_tool("inspect_ppt", inspect_ppt, str(pptx))
    for slide_index in args.inspect_slide:
        results["focused_slides"][str(slide_index)] = run_tool("inspect_slide", inspect_slide, str(pptx), slide_index)
    results["tools"]["check_ppt_quality"] = run_tool(
        "check_ppt_quality",
        check_ppt_quality,
        str(pptx),
        args.expected_title,
        args.required_number,
        str(out / "quality_report.md"),
    )
    results["after_hash"] = sha256(pptx)
    results["original_file_unchanged"] = results["before_hash"] == results["after_hash"]
    (out / "regression_results.json").write_text(json.dumps(results, ensure_ascii=False, indent=2), encoding="utf-8")
    return 0 if all(item.get("success") for item in results["tools"].values()) else 1


if __name__ == "__main__":
    raise SystemExit(main())
