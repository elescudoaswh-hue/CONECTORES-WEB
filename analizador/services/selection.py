"""Identidad estable de cada conector; el mismo CON puede tener varios nombres."""
from __future__ import annotations

import hashlib
import json
from analizador.core.engine import ConnectorRecord


def connector_selection_key(group: str, connector: ConnectorRecord) -> str:
    identity = (group, connector.name, connector.component, connector.kind)
    return hashlib.sha256(
        json.dumps(identity, ensure_ascii=False).encode("utf-8")
    ).hexdigest()
