"""Gestión de conexiones a SQL Server."""

from __future__ import annotations

import configparser
from pathlib import Path

import pyodbc


class Database:
    """Crea conexiones usando la configuración del entorno seleccionado."""

    def __init__(self, config_path: str | Path) -> None:
        self.config_path = Path(config_path)
        self.config = configparser.ConfigParser()
        loaded = self.config.read(self.config_path, encoding="utf-8")
        if not loaded:
            raise FileNotFoundError(
                f"No se pudo leer el archivo de configuración: {self.config_path}"
            )

    def connect(self) -> pyodbc.Connection:
        section = self.config["database"]
        driver = section.get("driver", "ODBC Driver 18 for SQL Server")
        server = section["server"]
        database = section["database"]
        trusted = section.getboolean("trusted_connection", fallback=True)
        encrypt = section.get("encrypt", "yes")
        trust_certificate = section.get("trust_server_certificate", "yes")

        parts = [
            f"DRIVER={{{driver}}}",
            f"SERVER={server}",
            f"DATABASE={database}",
            f"Encrypt={encrypt}",
            f"TrustServerCertificate={trust_certificate}",
            "Connection Timeout=5",
        ]
        if trusted:
            parts.append("Trusted_Connection=yes")
        else:
            parts.extend([
                f"UID={section['username']}",
                f"PWD={section['password']}",
            ])

        return pyodbc.connect(";".join(parts) + ";", timeout=5)
