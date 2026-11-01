"""Persistencia simple (SQLite) para el Cotizador de Pisos Blandos.

Idéntico en estructura a db.py del cotizador de pinturas: en Render (u otro
host con disco efímero) el archivo .db se pierde en cada deploy. Mientras no
haga falta persistencia real entre deploys, SQLite local es suficiente.
"""

import json
import os
import sqlite3
from datetime import datetime
from typing import Optional

DATA_DIR = os.environ.get("DATA_DIR", ".")
DB_PATH = os.path.join(DATA_DIR, "pisos_blandos.db")


def _conn():
    os.makedirs(DATA_DIR, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = _conn()
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS cotizaciones (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            fecha TEXT NOT NULL,
            cliente TEXT,
            vendedor TEXT,
            m2_total REAL,
            sectores_json TEXT NOT NULL,
            resultado_json TEXT NOT NULL,
            modificado_de INTEGER
        )
        """
    )
    try:
        conn.execute("ALTER TABLE cotizaciones ADD COLUMN modificado_de INTEGER")
    except sqlite3.OperationalError:
        pass
    conn.commit()
    conn.close()


def guardar_cotizacion(
    cliente: str, vendedor: str, m2_total: float, sectores: list, resultado: dict,
    modificado_de: Optional[int] = None,
) -> int:
    """Modificar una cotización NUNCA sobreescribe la original — siempre crea
    un N° nuevo con `modificado_de` apuntando al anterior."""
    conn = _conn()
    cur = conn.execute(
        """
        INSERT INTO cotizaciones
            (fecha, cliente, vendedor, m2_total, sectores_json, resultado_json, modificado_de)
        VALUES (?, ?, ?, ?, ?, ?, ?)
        """,
        (
            datetime.now().isoformat(timespec="seconds"),
            cliente,
            vendedor,
            m2_total,
            json.dumps(sectores, ensure_ascii=False),
            json.dumps(resultado, ensure_ascii=False),
            modificado_de,
        ),
    )
    conn.commit()
    numero = cur.lastrowid
    conn.close()
    return numero


def listar_cotizaciones(limit: int = 500) -> list:
    conn = _conn()
    filas = conn.execute(
        "SELECT id, fecha, cliente, vendedor, m2_total FROM cotizaciones ORDER BY id DESC LIMIT ?",
        (limit,),
    ).fetchall()
    conn.close()
    return [dict(f) for f in filas]


def obtener_cotizacion(numero: int) -> Optional[dict]:
    conn = _conn()
    fila = conn.execute("SELECT * FROM cotizaciones WHERE id = ?", (numero,)).fetchone()
    conn.close()
    if fila is None:
        return None
    d = dict(fila)
    d["sectores"] = json.loads(d.pop("sectores_json"))
    d["resultado"] = json.loads(d.pop("resultado_json"))
    return d
