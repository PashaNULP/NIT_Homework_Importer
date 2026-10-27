"""Конфігурація і абстракції збереження сесій NIT².

Модуль не залежить від macOS і дозволяє пізніше підключити Windows Credential
Manager без змін бізнес-логіки.
"""

from __future__ import annotations

import os
from dataclasses import dataclass
from typing import Any, Callable, Protocol


class SessionStore(Protocol):
    def load(self) -> dict[str, Any] | None: ...
    def save(self, session: dict[str, Any]) -> None: ...
    def delete(self) -> bool: ...


@dataclass(frozen=True)
class NitUserConfig:
    school_id: str
    user_id: str
    profile: str = "default"

    @classmethod
    def from_environment(cls, default_school_id: str, default_user_id: str) -> "NitUserConfig":
        return cls(
            school_id=os.getenv("NIT_SCHOOL_ID", default_school_id).strip(),
            user_id=os.getenv("NIT_USER_ID", default_user_id).strip(),
            profile=os.getenv("NIT_PROFILE", "default").strip() or "default",
        )


class FunctionSessionStore:
    """Сумісний адаптер над чинним Keychain-кодом."""

    def __init__(
        self,
        load: Callable[[], dict[str, Any] | None],
        save: Callable[[dict[str, Any]], None],
        delete: Callable[[], bool],
    ) -> None:
        self._load = load
        self._save = save
        self._delete = delete

    def load(self) -> dict[str, Any] | None:
        return self._load()

    def save(self, session: dict[str, Any]) -> None:
        self._save(session)

    def delete(self) -> bool:
        return self._delete()
