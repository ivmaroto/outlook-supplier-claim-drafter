"""Consultas de lectura y normalización de resultados SQL."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from .database import Database


class SupplierRepository:
    """Lee consultas externas y normaliza columnas según el mapping del entorno."""

    def __init__(
        self,
        database: Database,
        queries_dir: str | Path,
        mapping_path: str | Path,
    ) -> None:
        self.database = database
        self.queries_dir = Path(queries_dir)
        with Path(mapping_path).open("r", encoding="utf-8-sig") as file:
            self.mapping = json.load(file)

    def _read_query(self, filename: str) -> str:
        path = self.queries_dir / filename
        if not path.is_file():
            raise FileNotFoundError(f"No se encuentra la consulta SQL: {path}")
        return path.read_text(encoding="utf-8-sig").strip().rstrip(";")

    def _execute(self, filename: str, parameters: tuple[Any, ...] = ()) -> list[dict[str, Any]]:
        query = self._read_query(filename)
        with self.database.connect() as connection:
            cursor = connection.cursor()
            cursor.execute(query, parameters)
            columns = [column[0] for column in cursor.description]
            return [dict(zip(columns, row)) for row in cursor.fetchall()]

    @staticmethod
    def _normalize(
        rows: list[dict[str, Any]], mapping: dict[str, str]
    ) -> list[dict[str, Any]]:
        normalized = []
        for row in rows:
            normalized.append({
                internal_name: row.get(source_column)
                for internal_name, source_column in mapping.items()
            })
        return normalized

    def get_contacts(self) -> list[dict[str, Any]]:
        return self._normalize(
            self._execute("contactos.sql"), self.mapping["contacts"]
        )

    def get_suppliers(self) -> list[dict[str, Any]]:
        return self._normalize(
            self._execute("proveedores.sql"), self.mapping["suppliers"]
        )

    def get_order_lines(self, account: str) -> list[dict[str, Any]]:
        return self._normalize(
            self._execute("lineas_pedido.sql", (account,)),
            self.mapping["order_lines"],
        )
