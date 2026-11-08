"""Punto de entrada de desarrollo."""

from pathlib import Path
import sys
import tkinter as tk

ROOT_DIR = Path(__file__).resolve().parent
SRC_DIR = ROOT_DIR / "src"

sys.path.insert(0, str(SRC_DIR))

from outlook_supplier_claim_drafter.data.database import Database
from outlook_supplier_claim_drafter.data.repository import SupplierRepository
from outlook_supplier_claim_drafter.ui.main_window import MainWindow


CONFIG_PATH = ROOT_DIR / "config" / "environments" / "development.ini"
QUERIES_DIR = ROOT_DIR / "config" / "queries" / "development"
MAPPING_PATH = ROOT_DIR / "config" / "mappings" / "development.json"

root = tk.Tk()
repository = SupplierRepository(
    Database(CONFIG_PATH),
    QUERIES_DIR,
    MAPPING_PATH,
)
app = MainWindow(root, repository=repository)
root.mainloop()