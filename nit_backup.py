"""Надійні локальні JSON-backup для операцій запису NIT²."""

from __future__ import annotations

import hashlib
import json
import os
import secrets
import tempfile
from datetime import datetime
from pathlib import Path
from typing import Any


BACKUP_FORMAT = "NIT2-BACKUP-1"


def _canonical(data: Any) -> bytes:
    return json.dumps(data, ensure_ascii=False, sort_keys=True, separators=(",", ":")).encode("utf-8")


def write_secure_backup(payload: dict[str, Any], directory: Path, prefix: str) -> Path:
    """Атомарно створює унікальний backup з SHA-256 та правами лише власника."""
    directory.mkdir(parents=True, exist_ok=True, mode=0o700)
    try:
        os.chmod(directory, 0o700)
    except OSError:
        pass

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    path = directory / f"{prefix}_{stamp}_{secrets.token_hex(3)}.json"
    document = dict(payload)
    document["_backup"] = {
        "format": BACKUP_FORMAT,
        "sha256": hashlib.sha256(_canonical(payload)).hexdigest(),
    }

    fd, temporary = tempfile.mkstemp(prefix=".nit2_", suffix=".tmp", dir=directory)
    temporary_path = Path(temporary)
    try:
        os.fchmod(fd, 0o600)
        with os.fdopen(fd, "w", encoding="utf-8") as stream:
            json.dump(document, stream, ensure_ascii=False, indent=2)
            stream.write("\n")
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temporary_path, path)
        os.chmod(path, 0o600)
        verify_backup(path)
        return path
    except Exception:
        try:
            temporary_path.unlink(missing_ok=True)
        except OSError:
            pass
        raise


def verify_backup(path: Path) -> dict[str, Any]:
    """Перевіряє формат та контрольну суму, повертає корисне навантаження."""
    document = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(document, dict):
        raise ValueError("Backup не є JSON-об’єктом")
    metadata = document.pop("_backup", None)
    if not isinstance(metadata, dict) or metadata.get("format") != BACKUP_FORMAT:
        raise ValueError("Невідомий або застарілий формат backup")
    expected = str(metadata.get("sha256") or "")
    actual = hashlib.sha256(_canonical(document)).hexdigest()
    if not secrets.compare_digest(expected, actual):
        raise ValueError("Контрольна сума backup не збігається")
    return document
