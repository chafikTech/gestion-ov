#!/usr/bin/env python3
import argparse
import json
import os
import sys
from typing import Any, Dict


PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)


def emit_json(payload: Dict[str, Any]) -> None:
    sys.stdout.write(json.dumps(payload, ensure_ascii=False))
    sys.stdout.flush()


def parse_request() -> Dict[str, Any]:
    raw = sys.stdin.buffer.read()
    if not raw:
        return {}
    try:
        return json.loads(raw.decode("utf-8"))
    except Exception as exc:  # pragma: no cover
        raise ValueError(f"Invalid JSON input: {exc}") from exc


def normalize_script_name(script_value: Any) -> str:
    script = str(script_value or "").strip().replace("\\", "/")
    script = os.path.basename(script).lower()
    return script


def run_generate_document(payload: Dict[str, Any]) -> Dict[str, Any]:
    from src.python import generate_document

    document_type = str(payload.get("documentType") or "").strip()
    output_dir = str(payload.get("outputDir") or "").strip()

    if not output_dir:
        raise ValueError("Missing required field: outputDir")
    if not document_type:
        raise ValueError("Missing required field: documentType")

    generate_document.ensure_dir(output_dir)
    docx_name = generate_document.build_docx_filename(document_type, payload)
    docx_path = os.path.join(output_dir, docx_name)
    generate_document.generate_generic_docx({**payload, "documentType": document_type}, docx_path)

    return {"success": True, "docxFileName": docx_name, "docxFilePath": docx_path}


def run_generate_role(payload: Dict[str, Any]) -> Dict[str, Any]:
    from src.python import generate_role

    output_dir = str(payload.get("outputDir") or "").strip()
    if not output_dir:
        raise ValueError("Missing required field: outputDir")

    generate_role.ensure_dir(output_dir)
    safe_start = generate_role.safe_filename_part(payload.get("safeStart") or payload.get("periodStart") or "")
    safe_end = generate_role.safe_filename_part(payload.get("safeEnd") or payload.get("periodEnd") or "")

    docx_name = f"role_journees_{safe_start}_{safe_end}.docx"
    docx_path = os.path.join(output_dir, docx_name)
    generate_role.draw_role_docx(payload, docx_path)

    return {"success": True, "docxFileName": docx_name, "docxFilePath": docx_path}


def handle_generate(script: Any, payload: Any) -> Dict[str, Any]:
    script_name = normalize_script_name(script)
    data = payload if isinstance(payload, dict) else {}

    if script_name in ("generate_document.py", "generate_document"):
        return run_generate_document(data)
    if script_name in ("generate_role.py", "generate_role"):
        return run_generate_role(data)

    raise ValueError(f"Unsupported backend script: {script_name or '(empty)'}")


def main() -> None:
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("--ping", action="store_true")
    args, _ = parser.parse_known_args()

    if args.ping:
        emit_json({"success": True, "status": "ok"})
        return

    request = parse_request()
    command = str(request.get("command") or "generate").strip().lower()

    if command == "ping":
        emit_json({"success": True, "status": "ok"})
        return
    if command != "generate":
        raise ValueError(f"Unsupported backend command: {command}")

    result = handle_generate(request.get("script"), request.get("payload"))
    emit_json(result)


if __name__ == "__main__":
    try:
        main()
    except ModuleNotFoundError as exc:
        emit_json(
            {
                "success": False,
                "message": (
                    "Missing Python dependency for backend executable. "
                    "Rebuild with PyInstaller after installing backend/requirements.txt."
                ),
                "error": str(exc),
            }
        )
        sys.exit(1)
    except Exception as exc:  # pragma: no cover
        emit_json({"success": False, "message": str(exc)})
        sys.exit(1)
