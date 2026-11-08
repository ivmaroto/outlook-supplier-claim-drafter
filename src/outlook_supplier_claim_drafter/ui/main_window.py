"""Ventana principal Tkinter del Generador de Borradores Outlook."""

from __future__ import annotations

import json
import re
import html
import base64
import ctypes
import sys
import shutil
import webbrowser
from datetime import datetime, date
from html.parser import HTMLParser
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox, colorchooser, filedialog
import tkinter.font as tkfont
from typing import Any, Iterable, Mapping

from outlook_supplier_claim_drafter.data.repository import SupplierRepository


class RichTextEditor:
    """Editor visual Tkinter con formato básico y exportación HTML."""

    def __init__(self, parent, height=12):
        self.frame = ttk.Frame(parent)
        toolbar = ttk.Frame(self.frame)
        toolbar.pack(fill=tk.X, pady=(0, 5))

        self.font_var = tk.StringVar(value="Arial")
        self.size_var = tk.StringVar(value="11")

        self.font_box = ttk.Combobox(
            toolbar, textvariable=self.font_var,
            values=("Arial", "Calibri", "Times New Roman", "Verdana", "Tahoma"),
            width=16, state="readonly"
        )
        self.font_box.pack(side=tk.LEFT, padx=(0, 5))
        self.font_box.bind("<<ComboboxSelected>>", self._apply_font)

        self.size_box = ttk.Combobox(
            toolbar, textvariable=self.size_var,
            values=("8", "9", "10", "11", "12", "14", "16", "18", "20", "24", "28", "36"),
            width=4, state="readonly"
        )
        self.size_box.pack(side=tk.LEFT, padx=(0, 5))
        self.size_box.bind("<<ComboboxSelected>>", self._apply_size)

        ttk.Button(toolbar, text="B", width=3, command=lambda: self._toggle_tag("bold")).pack(side=tk.LEFT, padx=1)
        ttk.Button(toolbar, text="I", width=3, command=lambda: self._toggle_tag("italic")).pack(side=tk.LEFT, padx=1)
        ttk.Button(toolbar, text="U", width=3, command=lambda: self._toggle_tag("underline")).pack(side=tk.LEFT, padx=1)
        ttk.Button(toolbar, text="A Color", command=self._choose_foreground).pack(side=tk.LEFT, padx=3)
        ttk.Button(toolbar, text="Fondo", command=self._choose_background).pack(side=tk.LEFT, padx=1)

        body = ttk.Frame(self.frame)
        body.pack(fill=tk.BOTH, expand=True)
        self.text = tk.Text(body, wrap=tk.WORD, height=height, undo=True)
        scroll = ttk.Scrollbar(body, orient=tk.VERTICAL, command=self.text.yview)
        self.text.configure(yscrollcommand=scroll.set)
        self.text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)

        self.text.tag_configure("bold", font=("Arial", 11, "bold"))
        self.text.tag_configure("italic", font=("Arial", 11, "italic"))
        self.text.tag_configure("underline", underline=True)
        self._image_refs = []
        self._embedded_images = {}
        self.text.bind("<Control-v>", self._paste_clipboard)
        self.text.bind("<Shift-Insert>", self._paste_clipboard)

    def pack(self, **kwargs):
        self.frame.pack(**kwargs)

    def _selection(self):
        try:
            return self.text.index("sel.first"), self.text.index("sel.last")
        except tk.TclError:
            return None

    def _toggle_tag(self, tag):
        selection = self._selection()
        if not selection:
            return
        start, end = selection
        ranges = self.text.tag_ranges(tag)
        fully_applied = any(
            self.text.compare(a, "<=", start) and self.text.compare(b, ">=", end)
            for a, b in zip(ranges[0::2], ranges[1::2])
        )
        if fully_applied:
            self.text.tag_remove(tag, start, end)
        else:
            self.text.tag_add(tag, start, end)
        self.text.focus_set()

    def _apply_font(self, _event=None):
        self._apply_style("font_family", self.font_var.get())

    def _apply_size(self, _event=None):
        try:
            size = int(self.size_var.get())
        except ValueError:
            return
        self._apply_style("font_size", size)

    def _apply_style(self, prefix, value):
        selection = self._selection()
        if not selection:
            return
        start, end = selection
        tag = f"{prefix}_{str(value).replace(' ', '_')}"
        if prefix == "font_family":
            self.text.tag_configure(tag, font=(value, 11))
        else:
            self.text.tag_configure(tag, font=("Arial", value))
        self.text.tag_add(tag, start, end)
        self.text.focus_set()

    def _choose_color(self, option):
        selection = self._selection()
        if not selection:
            return
        color = colorchooser.askcolor(parent=self.frame.winfo_toplevel())[1]
        if not color:
            return
        start, end = selection
        tag = f"{option}_{color.lstrip('#')}"
        if option == "fg":
            self.text.tag_configure(tag, foreground=color)
        else:
            self.text.tag_configure(tag, background=color)
        self.text.tag_add(tag, start, end)
        self.text.focus_set()

    def _choose_foreground(self):
        self._choose_color("fg")

    def _choose_background(self):
        self._choose_color("bg")

    def clear(self):
        self.text.delete("1.0", tk.END)
        self._image_refs.clear()
        self._embedded_images.clear()

    @staticmethod
    def _read_windows_html_clipboard():
        """Lee el formato CF_HTML del portapapeles de Windows, si existe."""
        if sys.platform != "win32":
            return None
        user32 = ctypes.windll.user32
        kernel32 = ctypes.windll.kernel32
        user32.RegisterClipboardFormatW.restype = ctypes.c_uint
        user32.OpenClipboard.argtypes = [ctypes.c_void_p]
        user32.GetClipboardData.argtypes = [ctypes.c_uint]
        user32.GetClipboardData.restype = ctypes.c_void_p
        CF_HTML = user32.RegisterClipboardFormatW("HTML Format")
        if not user32.OpenClipboard(None):
            return None
        handle = None
        try:
            handle = user32.GetClipboardData(CF_HTML)
            if not handle:
                return None
            kernel32.GlobalLock.argtypes = [ctypes.c_void_p]
            kernel32.GlobalLock.restype = ctypes.c_void_p
            kernel32.GlobalSize.argtypes = [ctypes.c_void_p]
            kernel32.GlobalSize.restype = ctypes.c_size_t
            kernel32.GlobalUnlock.argtypes = [ctypes.c_void_p]
            pointer = kernel32.GlobalLock(handle)
            if not pointer:
                return None
            try:
                raw = ctypes.string_at(pointer, kernel32.GlobalSize(handle))
            finally:
                kernel32.GlobalUnlock(handle)
            decoded = raw.decode("utf-8", errors="replace")
            start = re.search(r"StartFragment:(\d+)", decoded)
            end = re.search(r"EndFragment:(\d+)", decoded)
            if start and end:
                fragment = decoded[int(start.group(1)):int(end.group(1))]
                if fragment:
                    return fragment
            marker_start = decoded.find("<!--StartFragment-->")
            marker_end = decoded.find("<!--EndFragment-->")
            if marker_start >= 0 and marker_end > marker_start:
                return decoded[marker_start + len("<!--StartFragment-->"):marker_end]
            return decoded
        finally:
            user32.CloseClipboard()

    @staticmethod
    def _read_windows_file_drop_clipboard():
        """Devuelve rutas copiadas desde el Explorador mediante CF_HDROP."""
        if sys.platform != "win32":
            return []
        user32 = ctypes.windll.user32
        shell32 = ctypes.windll.shell32
        user32.OpenClipboard.argtypes = [ctypes.c_void_p]
        user32.GetClipboardData.argtypes = [ctypes.c_uint]
        user32.GetClipboardData.restype = ctypes.c_void_p
        shell32.DragQueryFileW.argtypes = [
            ctypes.c_void_p, ctypes.c_uint, ctypes.c_wchar_p, ctypes.c_uint
        ]
        shell32.DragQueryFileW.restype = ctypes.c_uint
        if not user32.OpenClipboard(None):
            return []
        try:
            hdrop = user32.GetClipboardData(15)  # CF_HDROP
            if not hdrop:
                return []
            count = shell32.DragQueryFileW(hdrop, 0xFFFFFFFF, None, 0)
            paths = []
            for index in range(count):
                length = shell32.DragQueryFileW(hdrop, index, None, 0)
                buffer = ctypes.create_unicode_buffer(length + 1)
                shell32.DragQueryFileW(hdrop, index, buffer, length + 1)
                paths.append(buffer.value)
            return paths
        finally:
            user32.CloseClipboard()

    def _insert_image_file(self, path):
        """Inserta una imagen local y la conserva incrustada en el HTML."""
        suffix = Path(path).suffix.lower()
        mime_by_suffix = {
            ".png": "image/png", ".gif": "image/gif",
            ".jpg": "image/jpeg", ".jpeg": "image/jpeg",
        }
        mime = mime_by_suffix.get(suffix)
        if not mime:
            return False
        try:
            with open(path, "rb") as image_file:
                raw = image_file.read()
            encoded = base64.b64encode(raw).decode("ascii")
            try:
                photo = tk.PhotoImage(data=encoded)
            except tk.TclError:
                # JPEG y otros formatos pueden requerir Pillow.
                try:
                    from PIL import Image, ImageTk
                    import io
                    photo = ImageTk.PhotoImage(Image.open(io.BytesIO(raw)))
                except Exception:
                    return False
            image_name = f"signature_image_{len(self._image_refs)}"
            self.text.image_create(tk.INSERT, image=photo, name=image_name)
            self._image_refs.append(photo)
            self._embedded_images[image_name] = (mime, encoded)
            return True
        except (OSError, tk.TclError):
            return False

    def _paste_clipboard(self, _event=None):
        """Prioriza HTML e imágenes/archivos del portapapeles antes del texto."""
        fragment = self._read_windows_html_clipboard()
        if fragment:
            try:
                self._insert_html_fragment(fragment)
                return "break"
            except Exception:
                # Si el HTML no se puede interpretar, intentamos otros formatos.
                pass
        dropped_paths = self._read_windows_file_drop_clipboard()
        image_paths = [
            path for path in dropped_paths
            if Path(path).suffix.lower() in (".png", ".gif", ".jpg", ".jpeg")
        ]
        if image_paths:
            inserted = False
            for path in image_paths:
                if inserted:
                    self.text.insert(tk.INSERT, "\\n")
                inserted = self._insert_image_file(path) or inserted
            if inserted:
                return "break"

        # Algunas aplicaciones/gestores entregan la ruta como texto plano
        # en lugar de publicar CF_HDROP. Si apunta a una imagen existente,
        # insertamos el archivo y evitamos pegar la ruta literal.
        try:
            clipboard_text = self.text.clipboard_get().strip().strip('\"')
        except tk.TclError:
            clipboard_text = ""
        if clipboard_text:
            candidate = Path(clipboard_text)
            if candidate.is_file() and candidate.suffix.lower() in (
                ".png", ".gif", ".jpg", ".jpeg"
            ):
                if self._insert_image_file(str(candidate)):
                    return "break"

        try:
            self.text.event_generate("<<Paste>>")
        except tk.TclError:
            pass
        return "break"

    def _insert_html_fragment(self, fragment):
        parser = _BasicHTMLToText(self)
        parser.feed(fragment)
        parser.close()

    def set_html(self, value):
        self.clear()
        value = value or ""
        if not re.search(r"<\s*(p|div|br|span|b|strong|i|em|u)\b", value, re.I):
            value = value.replace("\\r\\n", "\n").replace("\\n", "\n").replace("\\r", "\n")
            self.text.insert("1.0", value)
            return
        parser = _BasicHTMLToText(self)
        try:
            parser.feed(value)
        except Exception:
            self.clear()
            self.text.insert("1.0", re.sub(r"<[^>]+>", "", value))

    def get_html(self):
        end = self.text.index("end-1c")
        if self.text.compare(end, "==", "1.0"):
            return ""
        result = []
        line_count = int(self.text.index("end-1c").split(".")[0])
        for line_no in range(1, line_count + 1):
            line_start = f"{line_no}.0"
            line_end = f"{line_no}.end"
            if line_no > 1:
                result.append("<br>")
            col = 0
            last_col = int(self.text.index(line_end).split(".")[1])
            while col < last_col:
                idx = f"{line_no}.{col}"
                image_names = [item[1] for item in self.text.dump(idx, f"{idx}+1c", image=True)
                               if item[0] == "image"]
                if image_names:
                    image_name = image_names[0]
                    image_data = self._embedded_images.get(image_name)
                    if image_data:
                        mime, encoded = image_data
                        result.append(f'<img src="data:{mime};base64,{encoded}" />')
                    col += 1
                    continue
                char = self.text.get(idx, f"{idx}+1c")
                tags = self.text.tag_names(idx)
                styles = []
                families = [t[len("font_family_"):].replace("_", " ") for t in tags if t.startswith("font_family_")]
                sizes = [t[len("font_size_"):] for t in tags if t.startswith("font_size_")]
                fgs = [t[len("fg_"):] for t in tags if t.startswith("fg_")]
                bgs = [t[len("bg_"):] for t in tags if t.startswith("bg_")]
                css = []
                if families:
                    css.append("font-family:" + html.escape(families[-1]) )
                if sizes:
                    css.append("font-size:" + sizes[-1] + "pt")
                if fgs:
                    css.append("color:#" + fgs[-1])
                if bgs:
                    css.append("background-color:#" + bgs[-1])
                if css:
                    styles.append('<span style="' + ";".join(css) + '">')
                if "bold" in tags:
                    styles.append("<strong>")
                if "italic" in tags:
                    styles.append("<em>")
                if "underline" in tags:
                    styles.append("<u>")
                links = [t[len("link_"):] for t in tags if t.startswith("link_")]
                if links:
                    styles.append('<a href="' + html.escape(links[-1], quote=True) + '">')
                rendered = html.escape(char).replace(" ", "&nbsp;")
                result.append("".join(styles) + rendered)
                if links:
                    result.append("</a>")
                if "underline" in tags:
                    result.append("</u>")
                if "italic" in tags:
                    result.append("</em>")
                if "bold" in tags:
                    result.append("</strong>")
                if css:
                    result.append("</span>")
                col += 1
        return "<html><body><p>" + "".join(result) + "</p></body></html>"


class _BasicHTMLToText(HTMLParser):
    """Importa el subconjunto HTML admitido por el editor."""
    def __init__(self, editor):
        from html.parser import HTMLParser as _HTMLParser
        self.editor = editor
        self.stack = []
        self._parser = _HTMLParser
        super().__init__()

    def handle_starttag(self, tag, attrs):
        attrs = dict(attrs)
        if tag == "br":
            self.editor.text.insert(tk.END, "\n")
            return
        if tag in ("p", "div") and self.editor.text.get("1.0", "end-1c"):
            self.editor.text.insert(tk.END, "\n")
        added = []
        if tag in ("b", "strong"):
            added.append("bold")
        elif tag in ("i", "em"):
            added.append("italic")
        elif tag == "u":
            added.append("underline")
        elif tag == "a":
            href = attrs.get("href", "")
            if href:
                name = "link_" + href
                self.editor.text.tag_configure(name, foreground="#0563C1", underline=True)
                added.append(name)
        elif tag in ("span", "p", "div"):
            style = attrs.get("style", "")
            for declaration in style.split(";"):
                if ":" not in declaration:
                    continue
                key, value = [part.strip() for part in declaration.split(":", 1)]
                if key == "color" and value.startswith("#"):
                    name = "fg_" + value[1:]
                    self.editor.text.tag_configure(name, foreground=value)
                    added.append(name)
                elif key == "background-color" and value.startswith("#"):
                    name = "bg_" + value[1:]
                    self.editor.text.tag_configure(name, background=value)
                    added.append(name)
                elif key == "font-size" and value.endswith("pt"):
                    name = "font_size_" + value[:-2]
                    self.editor.text.tag_configure(name, font=("Arial", int(float(value[:-2]))))
                    added.append(name)
                elif key == "font-family":
                    family = value.strip("'\\\"")
                    name = "font_family_" + family.replace(" ", "_")
                    self.editor.text.tag_configure(name, font=(family, 11))
                    added.append(name)
        if tag == "img":
            source = attrs.get("src", "")
            match = re.match(r"data:(image/(?:png|gif));base64,(.+)", source, re.I | re.S)
            if match:
                mime, encoded = match.group(1).lower(), re.sub(r"\s+", "", match.group(2))
                try:
                    raw = base64.b64decode(encoded, validate=True)
                    photo = tk.PhotoImage(data=base64.b64encode(raw).decode("ascii"))
                    image_name = f"signature_image_{len(self.editor._image_refs)}"
                    self.editor.text.image_create(tk.INSERT, image=photo, name=image_name)
                    self.editor._image_refs.append(photo)
                    self.editor._embedded_images[image_name] = (mime, encoded)
                except Exception:
                    self.editor.text.insert(tk.INSERT, "[Imagen no compatible]")
            else:
                self.editor.text.insert(tk.INSERT, "[Imagen externa no incrustada]")
            return
        self.stack.append((tag, added, self.editor.text.index(tk.INSERT)))

    def handle_endtag(self, tag):
        for index in range(len(self.stack) - 1, -1, -1):
            open_tag, added, start = self.stack[index]
            if open_tag == tag:
                end = self.editor.text.index(tk.INSERT)
                for name in added:
                    self.editor.text.tag_add(name, start, end)
                self.stack = self.stack[:index]
                if tag in ("p", "div"):
                    self.editor.text.insert(tk.END, "\n")
                break

    def handle_data(self, data):
        # Normaliza secuencias escapadas heredadas de plantillas anteriores.
        data = data.replace("\\r\\n", "\n").replace("\\n", "\n").replace("\\r", "\n")
        self.editor.text.insert(tk.END, data)


class MainWindow:

    def __init__(
        self,
        root: tk.Tk,
        repository: SupplierRepository | None = None,
    ) -> None:

        self.root = root
        self.repository = repository

        self.root.title("Generador de borradores Outlook")
        self.root.geometry("1200x760")
        self.root.minsize(900, 600)

        self._providers = []
        self._lines_by_provider = {}
        self._provider_draft_enabled = {}
        self._line_selected = {}
        self._line_selection_by_provider = {}
        self._line_column_selection_by_provider = {}
        self._bold_font = tkfont.nametofont("TkDefaultFont").copy()
        self._bold_font.configure(weight="bold")
        self._selected_provider = None
        self._provider_sort_column = "name"
        self._provider_sort_reverse = False
        self._contacts_by_account = {}
        self._email_dirty = False
        self._provider_email_session = {}
        self._email_templates_path = (
            Path(__file__).resolve().parents[3]
            / "config"
            / "email_templates.json"
        )
        self._email_templates = self._load_email_templates()
        self._signature_path = self._email_templates_path.parent / "email_signature.html"
        self._signature_dir = self._email_templates_path.parent / "email_signature"
        self._signature_manifest_path = self._signature_dir / "signature.json"
        self._signature_source_path = None
        self._signature_html = self._load_signature()

        self._build_layout()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        if self.repository is not None:
            self.load_initial_data()

    def _build_layout(self):

        notebook = ttk.Notebook(self.root)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Primera pestaña: trabajo diario con proveedores y pedidos.
        work_tab = ttk.Frame(notebook, padding=10)
        notebook.add(work_tab, text="Proveedores y pedidos")

        # Segunda pestaña: configuración persistente de la plantilla base.
        default_tab = ttk.Frame(notebook, padding=10)
        notebook.add(default_tab, text="Plantilla predeterminada")
        self._build_default_email_panel(default_tab)

        signature_tab = ttk.Frame(notebook, padding=10)
        notebook.add(signature_tab, text="Firma del remitente")
        self._build_signature_panel(signature_tab)

        # Al iniciar, mostrar directamente proveedores y pedidos.
        notebook.select(work_tab)

        upper = ttk.Panedwindow(
            work_tab,
            orient=tk.HORIZONTAL
        )
        upper.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        providers_frame = ttk.LabelFrame(
            upper,
            text="Proveedores",
            padding=8
        )

        email_frame = ttk.LabelFrame(
            upper,
            text="Correo del proveedor",
            padding=8
        )

        upper.add(providers_frame, weight=3)
        upper.add(email_frame, weight=2)

        self._build_providers_panel(providers_frame)
        self._build_email_panel(email_frame)

        lines_frame = ttk.LabelFrame(
            work_tab,
            text="Líneas del proveedor seleccionado",
            padding=8
        )
        lines_frame.pack(fill=tk.BOTH, expand=True)

        self._build_lines_panel(lines_frame)

        actions = ttk.Frame(work_tab)
        actions.pack(fill=tk.X, pady=(10, 0))

        self.generate_button = ttk.Button(
            actions,
            text="Generar borradores",
            state=tk.DISABLED
        )
        self.generate_button.pack(side=tk.RIGHT)

    def _build_providers_panel(self, parent):

        search_row = ttk.Frame(parent)
        search_row.pack(fill=tk.X, pady=(0, 8))

        ttk.Label(
            search_row,
            text="Buscar:"
        ).pack(side=tk.LEFT)

        self.search_var = tk.StringVar()

        entry = ttk.Entry(
            search_row,
            textvariable=self.search_var
        )

        entry.pack(
            side=tk.LEFT,
            fill=tk.X,
            expand=True,
            padx=(6, 0)
        )

        self.search_var.trace_add(
            "write",
            self._on_search_changed
        )

        columns = (
            "include",
            "account",
            "name",
            "orders"
        )

        self.providers_tree = ttk.Treeview(
            parent,
            columns=columns,
            show="headings",
            selectmode="browse"
        )

        labels = {
            "include": "Incluir",
            "account": "Cuenta",
            "name": "Proveedor",
            "orders": "Pedidos"
        }

        widths = {
            "include": 65,
            "account": 105,
            "name": 210,
            "orders": 75
        }

        for col in columns:

            self.providers_tree.heading(
                col,
                text=labels[col],
                command=lambda c=col: self._sort_providers(c)
            )

            self.providers_tree.column(
                col,
                width=widths[col],
                minwidth=55,
                anchor=(
                    tk.CENTER
                    if col in ("include", "orders")
                    else tk.W
                )
            )

        scroll = ttk.Scrollbar(
            parent,
            orient=tk.VERTICAL,
            command=self.providers_tree.yview
        )

        self.providers_tree.configure(
            yscrollcommand=scroll.set
        )
        self.providers_tree.tag_configure("selected_bold", font=self._bold_font)

        self.providers_tree.pack(
            side=tk.LEFT,
            fill=tk.BOTH,
            expand=True
        )

        scroll.pack(side=tk.RIGHT, fill=tk.Y)

        self.providers_tree.bind(
            "<<TreeviewSelect>>",
            self._on_provider_selected
        )

        self.providers_tree.bind(
            "<Button-1>",
            self._on_provider_click
        )

    def _build_email_panel(self, parent):

        # Editor del correo individual del proveedor seleccionado.
        ttk.Label(parent, text="Asunto").pack(anchor=tk.W)
        self.subject_entry = ttk.Entry(parent)
        self.subject_entry.pack(fill=tk.X, pady=(4, 10))

        ttk.Label(parent, text="Mensaje").pack(anchor=tk.W)
        self.body_editor = RichTextEditor(parent, height=12)
        self.body_editor.pack(fill=tk.BOTH, expand=True, pady=(4, 0))

        self.email_provider_label = ttk.Label(
            parent,
            text="Mensaje para este envío (los cambios no se guardan)."
        )
        self.email_provider_label.pack(anchor=tk.W, pady=(8, 0))

    def _build_default_email_panel(self, parent):

        ttk.Label(parent, text="Asunto predeterminado").pack(anchor=tk.W)
        self.default_subject_entry = ttk.Entry(parent)
        self.default_subject_entry.pack(fill=tk.X, pady=(4, 10))

        ttk.Label(parent, text="Mensaje predeterminado").pack(anchor=tk.W)
        self.default_body_editor = RichTextEditor(parent, height=12)
        self.default_body_editor.pack(fill=tk.BOTH, expand=True, pady=(4, 0))

        default_actions = ttk.Frame(parent)
        default_actions.pack(fill=tk.X, pady=(8, 0))
        ttk.Button(
            default_actions,
            text="Guardar plantilla predeterminada",
            command=self._save_default_email
        ).pack(side=tk.LEFT)

        self._populate_default_email_fields()

    def _build_signature_panel(self, parent):
        ttk.Label(
            parent,
            text="Adjunta la firma exportada desde Outlook (.htm o .html). "
                 "Se conservará también su carpeta de imágenes."
        ).pack(anchor=tk.W, pady=(0, 8))

        file_frame = ttk.Frame(parent)
        file_frame.pack(fill=tk.X, pady=(0, 8))

        self.signature_file_label = ttk.Label(
            file_frame,
            text=self._signature_source_path.name
            if self._signature_source_path else "No hay ninguna firma cargada.",
            anchor=tk.W
        )
        self.signature_file_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

        ttk.Button(
            file_frame,
            text="Adjuntar firma",
            command=self._import_signature
        ).pack(side=tk.RIGHT, padx=(8, 0))

        ttk.Button(
            file_frame,
            text="Visualizar",
            command=self._open_signature_preview
        ).pack(side=tk.RIGHT)

        ttk.Label(
            parent,
            text=(
                "La vista previa se abrirá en el navegador predeterminado. "
                "Para conservar las imágenes, mantén junto al archivo HTML "
                "la carpeta original «NombreFirma_files»."
            ),
            wraplength=900,
            justify=tk.LEFT
        ).pack(anchor=tk.W, pady=(0, 8))

        self.signature_status_label = ttk.Label(
            parent,
            text="Firma guardada." if self._signature_source_path else "Pendiente de importar una firma."
        )
        self.signature_status_label.pack(anchor=tk.W, pady=(0, 8))

        actions = ttk.Frame(parent)
        actions.pack(fill=tk.X, pady=(8, 0))
        ttk.Button(
            actions,
            text="Guardar firma",
            command=self._save_signature
        ).pack(side=tk.LEFT)

    def _load_signature(self):
        """Carga la firma importada mediante su manifiesto; admite el HTML antiguo."""
        try:
            if self._signature_manifest_path.exists():
                manifest = json.loads(
                    self._signature_manifest_path.read_text(encoding="utf-8")
                )
                filename = Path(manifest.get("filename", "")).name
                if not filename:
                    return ""
                source_path = self._signature_dir / filename
                if not source_path.is_file():
                    return ""
                self._signature_source_path = source_path
                return source_path.read_text(encoding="utf-8-sig")

            # Compatibilidad con la firma HTML guardada por versiones anteriores.
            if self._signature_path.is_file():
                self._signature_source_path = self._signature_path
                return self._signature_path.read_text(encoding="utf-8-sig")

            return ""
        except (OSError, ValueError, TypeError) as exc:
            messagebox.showwarning(
                "Firma del remitente",
                f"No se pudo leer la firma guardada:\\n{exc}",
                parent=self.root
            )
            return ""

    def _import_signature(self):
        selected = filedialog.askopenfilename(
            parent=self.root,
            title="Seleccionar firma exportada desde Outlook",
            filetypes=(
                ("Archivos HTML", "*.htm *.html"),
                ("Todos los archivos", "*.*")
            )
        )
        if not selected:
            return

        source_path = Path(selected)
        try:
            self._signature_dir.mkdir(parents=True, exist_ok=True)

            # Se conserva el nombre original para no romper las referencias
            # relativas a la carpeta «NombreFirma_files» de Outlook.
            destination = self._signature_dir / source_path.name
            shutil.copy2(source_path, destination)

            source_assets = source_path.with_name(source_path.stem + "_files")
            destination_assets = self._signature_dir / source_assets.name
            if destination_assets.exists():
                shutil.rmtree(destination_assets)
            if source_assets.is_dir():
                shutil.copytree(source_assets, destination_assets)

            self._signature_source_path = destination
            self._signature_html = destination.read_text(encoding="utf-8-sig")
            self.signature_file_label.configure(text=source_path.name)
            self.signature_status_label.configure(
                text="Firma importada. Pulsa «Guardar firma» para conservarla."
            )
        except (OSError, UnicodeError) as exc:
            messagebox.showerror(
                "Error al importar la firma",
                f"No se pudo importar la firma:\\n{exc}",
                parent=self.root
            )

    def _open_signature_preview(self):
        if not self._signature_source_path or not self._signature_source_path.is_file():
            messagebox.showinfo(
                "Vista previa",
                "Primero debes adjuntar una firma HTML.",
                parent=self.root
            )
            return

        try:
            webbrowser.open(self._signature_source_path.resolve().as_uri())
        except OSError as exc:
            messagebox.showerror(
                "Error al abrir la vista previa",
                str(exc),
                parent=self.root
            )

    def _save_signature(self):
        if not self._signature_source_path or not self._signature_source_path.is_file():
            messagebox.showwarning(
                "Firma del remitente",
                "Primero debes adjuntar un archivo de firma HTML.",
                parent=self.root
            )
            return False

        try:
            self._signature_dir.mkdir(parents=True, exist_ok=True)

            # Si procede de una versión anterior, copiarlo a la nueva
            # estructura antes de escribir el manifiesto.
            if self._signature_source_path.parent != self._signature_dir:
                migrated_path = self._signature_dir / self._signature_source_path.name
                shutil.copy2(self._signature_source_path, migrated_path)
                self._signature_source_path = migrated_path

            manifest = {"filename": self._signature_source_path.name}
            self._signature_manifest_path.write_text(
                json.dumps(manifest, ensure_ascii=False, indent=2),
                encoding="utf-8"
            )
            self.signature_status_label.configure(text="Firma guardada correctamente.")
            messagebox.showinfo(
                "Firma del remitente",
                "La firma se ha guardado correctamente.",
                parent=self.root
            )
            return True
        except OSError as exc:
            messagebox.showerror(
                "Error al guardar la firma",
                str(exc),
                parent=self.root
            )
            return False

    def _build_lines_panel(self, parent):
        selection_actions = ttk.Frame(parent)
        selection_actions.pack(fill=tk.X, pady=(0, 6))

        self.all_lines_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            selection_actions,
            text="✓",
            variable=self.all_lines_var,
            command=self._toggle_all_lines
        ).pack(side=tk.LEFT)

        self._line_columns = ("order", "product", "description", "quantity")
        self._line_column_labels = {
            "order": "Pedido",
            "product": "Producto",
            "description": "Descripción",
            "quantity": "Cantidad"
        }
        self._line_column_widths = {
            "order": 130,
            "product": 170,
            "description": 420,
            "quantity": 110
        }
        self._line_filter_vars = {
            col: tk.StringVar() for col in self._line_columns
        }
        self._line_filter_entries = {}
        self._line_display_columns = list(self._line_columns)
        self._line_filter_after_id = None
        self._line_sort_column = None
        self._line_sort_reverse = False

        # Las columnas empiezan seleccionadas para conservar el comportamiento
        # visual de la versión blanca: una fila incluida aparece blanca en todas.
        self._line_column_selected = {col: True for col in self._line_columns}
        self._visible_line_rows = []
        self._line_header_press = None
        self._line_resize_column = None
        self._line_resize_start_x = 0
        self._line_resize_start_width = 0
        self._last_focused_line_id = None

        table_frame = ttk.Frame(parent)
        table_frame.pack(fill=tk.BOTH, expand=True)
        table_frame.grid_rowconfigure(2, weight=1)
        table_frame.grid_columnconfigure(1, weight=1)

        self._line_filter_canvas = tk.Canvas(
            table_frame, height=29, highlightthickness=0, borderwidth=0,
            background="#f0f0f0"
        )
        self._line_filter_canvas.grid(row=0, column=1, sticky="ew")
        self._line_filter_canvas.bind(
            "<Configure>", lambda _event: self._layout_line_filter_entries()
        )

        # Encabezados fijos; los cuerpos tienen desplazamiento vertical.
        self.include_header_canvas = tk.Canvas(
            table_frame, width=75, height=34, highlightthickness=0,
            borderwidth=0, background="#eeeeee"
        )
        self.include_header_canvas.grid(row=1, column=0, sticky="ew")
        self.lines_header_canvas = tk.Canvas(
            table_frame, height=34, highlightthickness=0,
            borderwidth=0, background="#eeeeee"
        )
        self.lines_header_canvas.grid(row=1, column=1, sticky="ew")

        self.include_canvas = tk.Canvas(
            table_frame, width=75, highlightthickness=0, borderwidth=0,
            background="white"
        )
        self.include_canvas.grid(row=2, column=0, sticky="nsew")
        self.lines_canvas = tk.Canvas(
            table_frame, highlightthickness=0, borderwidth=0,
            background="white"
        )
        self.lines_canvas.grid(row=2, column=1, sticky="nsew")

        scroll_y = ttk.Scrollbar(table_frame, orient=tk.VERTICAL)
        scroll_y.grid(row=2, column=2, sticky="ns")
        scroll_x = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL)
        scroll_x.grid(row=3, column=1, sticky="ew")

        def sync_vertical(first, last):
            self.lines_canvas.yview_moveto(first)
            self.include_canvas.yview_moveto(first)
            scroll_y.set(first, last)

        def scroll_both(*args):
            self.lines_canvas.yview(*args)
            self.include_canvas.yview(*args)

        self.lines_canvas.configure(
            yscrollcommand=sync_vertical,
            xscrollcommand=self._on_lines_horizontal_scroll
        )
        self.include_canvas.configure(yscrollcommand=lambda _first, _last: None)
        scroll_y.configure(command=scroll_both)
        scroll_x.configure(command=self._on_lines_horizontal_scroll_command)
        self._lines_horizontal_scrollbar = scroll_x

        for col in self._line_columns:
            self._line_filter_vars[col].trace_add(
                "write", lambda *_args: self._schedule_line_filter()
            )

        self.lines_header_canvas.bind("<ButtonPress-1>", self._on_line_header_press)
        self.lines_header_canvas.bind("<B1-Motion>", self._on_line_header_motion)
        self.lines_header_canvas.bind("<ButtonRelease-1>", self._on_line_header_release)
        self.lines_header_canvas.bind("<Double-Button-1>", self._on_line_header_double_click)
        self.include_canvas.bind("<Button-1>", self._on_include_canvas_click)
        self.lines_canvas.bind("<Button-1>", self._on_lines_canvas_click)
        self.include_canvas.bind("<Return>", self._on_line_enter)
        self.lines_canvas.bind("<Return>", self._on_line_enter)
        for widget in (self.lines_canvas, self.include_canvas):
            widget.bind("<Up>", self._on_line_navigation)
            widget.bind("<Down>", self._on_line_navigation)
            widget.bind("<Home>", self._on_line_navigation)
            widget.bind("<End>", self._on_line_navigation)

        status_frame = ttk.Frame(parent)
        status_frame.pack(fill=tk.X, pady=(6, 0))
        self.lines_status = ttk.Label(
            status_frame,
            text="Selecciona un proveedor para ver sus líneas."
        )
        self.lines_status.pack(side=tk.LEFT, anchor=tk.W)

        self._rebuild_line_filter_row()
        self._draw_lines_table()

    def _line_total_width(self):
        return sum(self._line_column_widths[col] for col in self._line_display_columns)

    def _on_lines_horizontal_scroll(self, first, last):
        self._line_filter_canvas.xview_moveto(first)
        self.lines_header_canvas.xview_moveto(first)
        self._lines_horizontal_scrollbar.set(first, last)

    def _on_lines_horizontal_scroll_command(self, *args):
        self.lines_canvas.xview(*args)
        self.lines_header_canvas.xview(*args)
        self._line_filter_canvas.xview(*args)

    def _schedule_line_filter(self):
        if self._line_filter_after_id is not None:
            self.root.after_cancel(self._line_filter_after_id)
        self._line_filter_after_id = self.root.after(120, self._apply_line_filters)

    def _apply_line_filters(self):
        self._line_filter_after_id = None
        self._refresh_lines()

    def _sort_lines_by_column(self, column):
        if column not in self._line_columns:
            return
        if self._line_sort_column == column:
            self._line_sort_reverse = not self._line_sort_reverse
        else:
            self._line_sort_column = column
            self._line_sort_reverse = False
        self._refresh_lines()

    @staticmethod
    def _line_sort_key(value, column):
        if value is None:
            return (1, "")
        if column == "quantity":
            try:
                normalized = str(value).strip().replace(" ", "").replace(",", ".")
                return (0, float(normalized))
            except (TypeError, ValueError):
                return (1, str(value).casefold())
        return (0, str(value).casefold())

    def _line_column_x_positions(self):
        positions = {}
        x = 0
        for col in self._line_display_columns:
            positions[col] = (x, x + self._line_column_widths[col])
            x += self._line_column_widths[col]
        return positions

    def _line_column_at_x(self, x):
        for col, (left, right) in self._line_column_x_positions().items():
            if left <= x < right:
                return col, left, right
        return None, None, None

    def _line_resize_target_at_x(self, x, tolerance=7):
        """Devuelve la columna cuyo separador está bajo el puntero."""
        positions = self._line_column_x_positions()
        for col in self._line_display_columns:
            _left, right = positions[col]
            if abs(x - right) <= tolerance:
                return col
        return None

    def _auto_fit_line_column(self, col):
        """Ajusta una columna al encabezado y al contenido actualmente visible."""
        font = tkfont.nametofont("TkDefaultFont")
        bold_font = font.copy()
        bold_font.configure(weight="bold")
        label = self._line_column_labels[col]
        if self._line_sort_column == col:
            label += "  " + ("▼" if self._line_sort_reverse else "▲")

        # El encabezado reserva espacio para la casilla de selección.
        required_width = font.measure(label) + 42
        for _row_id, row in self._visible_line_rows:
            value = row.get(col, "")
            value = "" if value is None else str(value)
            required_width = max(
                required_width,
                (bold_font if self._line_selected.get(_row_id, True) else font).measure(value) + 18
            )

        self._line_column_widths[col] = max(65, required_width)
        self._layout_line_filter_entries()
        self._line_filter_canvas.configure(
            scrollregion=(0, 0, self._line_total_width(), 28)
        )
        self._draw_lines_table()

    def _on_line_header_double_click(self, event):
        if event.y > 34:
            return "break"
        x = self.lines_header_canvas.canvasx(event.x)
        col = self._line_resize_target_at_x(x)
        if col:
            self._auto_fit_line_column(col)
        return "break"

    def _on_line_header_press(self, event):
        x = self.lines_header_canvas.canvasx(event.x)
        y = event.y
        if y > 34:
            self._line_header_press = None
            return

        # Detectar primero los separadores: el punto exacto del borde también
        # pertenece geométricamente a la columna siguiente.
        resize_col = self._line_resize_target_at_x(x)
        if resize_col:
            self._line_resize_column = resize_col
            self._line_resize_start_x = event.x
            self._line_resize_start_width = self._line_column_widths[resize_col]
            self._line_header_press = None
            return

        col, left, right = self._line_column_at_x(x)
        if not col:
            self._line_header_press = None
            return

        self._line_resize_column = None
        self._line_header_press = {
            "column": col,
            "x": x,
            "y": y,
            "moved": False
        }

    def _on_line_header_motion(self, event):
        if self._line_resize_column:
            delta = event.x - self._line_resize_start_x
            col = self._line_resize_column
            self._line_column_widths[col] = max(
                65, self._line_resize_start_width + delta
            )
            self._layout_line_filter_entries()
            self._line_filter_canvas.configure(
                scrollregion=(0, 0, self._line_total_width(), 28)
            )
            self._draw_lines_table()
            return

        if self._line_header_press:
            current_x = self.lines_header_canvas.canvasx(event.x)
            if abs(current_x - self._line_header_press["x"]) > 8:
                self._line_header_press["moved"] = True

    def _on_line_header_release(self, event):
        if self._line_resize_column:
            self._line_resize_column = None
            return

        press = self._line_header_press
        self._line_header_press = None
        if not press:
            return

        x = self.lines_header_canvas.canvasx(event.x)
        y = event.y
        source = press["column"]

        if press["moved"] and y <= 34:
            target, _, _ = self._line_column_at_x(x)
            if target and target != source:
                columns = list(self._line_display_columns)
                columns.remove(source)
                columns.insert(columns.index(target), source)
                self._line_display_columns = columns
                self._rebuild_line_filter_row()
                self._draw_lines_table()
            return

        if y > 34:
            return

        left, right = self._line_column_x_positions()[source]
        # Casilla de selección de columna en el encabezado.
        if x <= left + 27:
            self._toggle_line_column_selection(source)
        else:
            self._sort_lines_by_column(source)

    def _toggle_line_column_selection(self, column):
        self._line_column_selected[column] = not self._line_column_selected.get(column, True)
        if self._selected_provider:
            self._line_column_selection_by_provider[self._selected_provider] = dict(
                self._line_column_selected
            )
        self._draw_lines_table()

    def _on_include_canvas_click(self, event):
        y = self.include_canvas.canvasy(event.y)
        row_index = int(y // 28)
        if row_index < 0 or row_index >= len(self._visible_line_rows):
            return "break"
        row_id, _row = self._visible_line_rows[row_index]
        self._last_focused_line_id = row_id
        self.include_canvas.focus_set()
        self._toggle_line_selection(row_id)
        return "break"

    def _on_lines_canvas_click(self, event):
        """Establece como activa la misma fila de datos que se ha pulsado."""
        y = self.lines_canvas.canvasy(event.y)
        row_index = int(y // 28)
        if 0 <= row_index < len(self._visible_line_rows):
            row_id, _row = self._visible_line_rows[row_index]
            self._last_focused_line_id = row_id
            self.lines_canvas.focus_set()
            self._draw_lines_table()
        return "break"

    def _rebuild_line_filter_row(self):
        self._line_filter_canvas.delete("all")
        self._line_filter_entries = {}
        x = 0
        for col in self._line_display_columns:
            width = self._line_column_widths[col]
            entry = ttk.Entry(
                self._line_filter_canvas,
                textvariable=self._line_filter_vars[col]
            )
            self._line_filter_entries[col] = entry
            self._line_filter_canvas.create_window(
                x + 2, 2, anchor="nw", window=entry,
                width=max(30, width - 4), height=24,
                tags=(f"filter_{col}",)
            )
            x += width
        self._line_filter_canvas.configure(scrollregion=(0, 0, x, 28))
        self._line_filter_canvas.xview_moveto(0)

    def _layout_line_filter_entries(self):
        if not hasattr(self, "lines_canvas") or not hasattr(self, "_line_filter_canvas"):
            return
        x = 0
        for col in self._line_display_columns:
            width = self._line_column_widths[col]
            windows = self._line_filter_canvas.find_withtag(f"filter_{col}")
            if windows:
                self._line_filter_canvas.coords(windows[0], x + 2, 2)
                self._line_filter_canvas.itemconfigure(
                    windows[0], width=max(30, width - 4)
                )
            x += width
        self._line_filter_canvas.configure(scrollregion=(0, 0, x, 28))

    def _toggle_line_selection(self, row_id):
        selected = not self._line_selected.get(row_id, True)
        self._line_selected[row_id] = selected
        if self._selected_provider:
            self._line_selection_by_provider[self._selected_provider] = dict(self._line_selected)
        visible_ids = [iid for iid, _row in self._visible_line_rows]
        self.all_lines_var.set(
            bool(visible_ids) and all(self._line_selected.get(iid, False) for iid in visible_ids)
        )
        self._draw_lines_table()

    def _draw_checkbox(self, canvas, x, y, checked, foreground="#222222"):
        canvas.create_rectangle(
            x, y, x + 13, y + 13, outline=foreground, fill="white", width=1
        )
        if checked:
            canvas.create_line(
                x + 2, y + 7, x + 5, y + 11, x + 11, y + 2,
                fill=foreground, width=2, capstyle=tk.ROUND, joinstyle=tk.ROUND
            )

    def _draw_lines_table(self):
        if not hasattr(self, "lines_canvas"):
            return

        header_height = 34
        row_height = 28
        include_width = 75

        self.lines_header_canvas.delete("all")
        self.include_header_canvas.delete("all")
        self.lines_canvas.delete("all")
        self.include_canvas.delete("all")

        # Encabezado fijo de inclusión.
        self.include_header_canvas.create_rectangle(
            0, 0, include_width, header_height,
            fill="#eeeeee", outline="#c8c8c8"
        )
        self.include_header_canvas.create_text(
            include_width / 2, header_height / 2,
            text="Incluir", anchor="center", fill="#222222"
        )

        positions = self._line_column_x_positions()
        for col in self._line_display_columns:
            left, right = positions[col]
            self.lines_header_canvas.create_rectangle(
                left, 0, right, header_height,
                fill="#eeeeee", outline="#c8c8c8"
            )
            self._draw_checkbox(
                self.lines_header_canvas, left + 7, 10,
                self._line_column_selected.get(col, True)
            )
            label = self._line_column_labels[col]
            if self._line_sort_column == col:
                label += "  " + ("▼" if self._line_sort_reverse else "▲")
            self.lines_header_canvas.create_text(
                left + 28, header_height / 2, text=label,
                anchor="w", fill="#222222"
            )

        for row_index, (row_id, row) in enumerate(self._visible_line_rows):
            y0 = row_index * row_height
            y1 = y0 + row_height
            row_selected = self._line_selected.get(row_id, True)

            self.include_canvas.create_rectangle(
                0, y0, include_width, y1,
                fill="#ffffff" if row_selected else "#e2e2e2",
                outline="#d0d0d0"
            )
            self._draw_checkbox(
                self.include_canvas, (include_width - 13) / 2,
                y0 + 7, row_selected
            )

            values = {
                "order": row.get("order", ""),
                "product": row.get("product", ""),
                "description": row.get("description", ""),
                "quantity": row.get("quantity", "")
            }
            for col in self._line_display_columns:
                left, right = positions[col]
                cell_selected = (
                    row_selected and self._line_column_selected.get(col, True)
                )
                background = "#ffffff" if cell_selected else "#e2e2e2"
                self.lines_canvas.create_rectangle(
                    left, y0, right, y1,
                    fill=background, outline="#d0d0d0"
                )
                raw_value = values.get(col, "")
                value = str(raw_value if raw_value is not None else "")
                max_chars = max(1, int((right - left - 16) / 7))
                if len(value) > max_chars:
                    value = value[:max_chars - 1] + "…"
                anchor = "center" if col == "quantity" else "w"
                text_x = (left + right) / 2 if col == "quantity" else left + 8
                self.lines_canvas.create_text(
                    text_x, (y0 + y1) / 2, text=value, anchor=anchor,
                    fill="#222222",
                    font=("TkDefaultFont", 9, "bold" if row_selected else "normal")
                )

            # Indicador de foco: permite localizar la fila al navegar con teclado.
            if row_id == self._last_focused_line_id:
                self.include_canvas.create_rectangle(
                    1, y0 + 1, include_width - 1, y1 - 1,
                    outline="#2878c8", width=2
                )
                self.lines_canvas.create_rectangle(
                    1, y0 + 1, self._line_total_width() - 1, y1 - 1,
                    outline="#2878c8", width=2
                )

        total_width = self._line_total_width()
        body_height = max(1, len(self._visible_line_rows) * row_height)
        self.lines_header_canvas.configure(
            scrollregion=(0, 0, total_width, header_height)
        )
        self.lines_canvas.configure(
            scrollregion=(0, 0, total_width, body_height)
        )
        self.include_header_canvas.configure(
            scrollregion=(0, 0, include_width, header_height)
        )
        self.include_canvas.configure(
            scrollregion=(0, 0, include_width, body_height)
        )
        self._line_filter_canvas.configure(
            scrollregion=(0, 0, total_width, 28)
        )

    @staticmethod
    def _match_text_filter(value, expression):
        expression = expression.strip().casefold()
        if not expression:
            return True
        text = str(value if value is not None else "").casefold()
        # Sin comodín: contiene. Con comodines: *texto*, texto* o *texto.
        if "*" not in expression:
            return expression in text
        pattern = "^" + re.escape(expression).replace(r"\*", ".*") + "$"
        return re.search(pattern, text) is not None

    @staticmethod
    def _match_numeric_filter(value, expression):
        expression = expression.strip().replace(",", ".")
        if not expression:
            return True
        try:
            number = float(str(value).replace(",", "."))
        except (TypeError, ValueError):
            return False
        range_match = re.fullmatch(r"\s*(-?\d+(?:\.\d+)?)\s*\.\.\s*(-?\d+(?:\.\d+)?)\s*", expression)
        if range_match:
            low, high = map(float, range_match.groups())
            return min(low, high) <= number <= max(low, high)
        match = re.fullmatch(r"(>=|<=|>|<|=)?\s*(-?\d+(?:\.\d+)?)", expression)
        if not match:
            return False
        operator, target = match.groups()
        target = float(target)
        return {
            None: number == target,
            "=": number == target,
            ">": number > target,
            ">=": number >= target,
            "<": number < target,
            "<=": number <= target
        }[operator]

    @staticmethod
    def _match_date_filter(value, expression):
        expression = expression.strip()
        if not expression:
            return True
        def parse_date(raw):
            raw = raw.strip()
            for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
                try:
                    return datetime.strptime(raw, fmt).date()
                except ValueError:
                    pass
            raise ValueError(raw)
        try:
            actual = value if isinstance(value, date) else parse_date(str(value))
            range_match = re.fullmatch(r"(.+?)\s*\.\.\s*(.+)", expression)
            if range_match:
                first, last = map(parse_date, range_match.groups())
                return min(first, last) <= actual <= max(first, last)
            match = re.fullmatch(r"(>=|<=|>|<|=)?\s*(.+)", expression)
            if not match:
                return False
            operator, target_raw = match.groups()
            target = parse_date(target_raw)
            return {
                None: actual == target,
                "=": actual == target,
                ">": actual > target,
                ">=": actual >= target,
                "<": actual < target,
                "<=": actual <= target
            }[operator]
        except (TypeError, ValueError):
            return False

    def _apply_line_filters(self):
        if not hasattr(self, "lines_canvas"):
            return
        self._refresh_lines()

    def _load_email_templates(self):
        empty = {"default": {"subject": "", "body": ""}}

        try:
            if not self._email_templates_path.is_file():
                return empty

            with self._email_templates_path.open("r", encoding="utf-8-sig") as file:
                data = json.load(file)

            if not isinstance(data, dict):
                return empty

            default_data = data.get("default", {})
            if not isinstance(default_data, dict):
                default_data = {}

            providers_data = data.get("providers", {})
            if not isinstance(providers_data, dict):
                providers_data = {}

            normalized_providers = {}
            for account, provider_data in providers_data.items():
                if isinstance(provider_data, dict):
                    normalized_providers[str(account)] = {
                        "subject": str(provider_data.get("subject", "")),
                        "body": str(provider_data.get("body", ""))
                    }

            return {
                "default": {
                    "subject": str(default_data.get("subject", "")),
                    "body": str(default_data.get("body", ""))
                },
                "providers": normalized_providers
            }
        except (OSError, json.JSONDecodeError, AttributeError, TypeError):
            messagebox.showwarning(
                "Plantilla de correo",
                "No se pudo leer email_templates.json. Se utilizará una plantilla vacía.",
                parent=self.root
            )
            return empty

    def _write_email_templates(self):
        self._email_templates_path.parent.mkdir(parents=True, exist_ok=True)
        with self._email_templates_path.open("w", encoding="utf-8") as file:
            json.dump(
                self._email_templates,
                file,
                ensure_ascii=False,
                indent=2
            )

    def _populate_default_email_fields(self):
        default = self._email_templates["default"]
        self.default_subject_entry.delete(0, tk.END)
        self.default_subject_entry.insert(0, default.get("subject", ""))
        self.default_body_editor.set_html(default.get("body", ""))

    def _save_default_email(self, show_confirmation=True):
        self._email_templates = {
            "default": {
            "subject": self.default_subject_entry.get(),
                "body": self.default_body_editor.get_html()
            }
        }
        try:
            self._write_email_templates()
        except OSError as exc:
            messagebox.showerror(
                "Error al guardar",
                f"No se pudo guardar la plantilla predeterminada:\\n{exc}",
                parent=self.root
            )
            return False

        # El editor de envío refleja la plantilla guardada solo cuando
        # todavía no hay proveedor seleccionado.
        if not self._selected_provider:
            self._load_default_into_provider_editor()

        if show_confirmation:
            messagebox.showinfo(
                "Plantilla guardada",
                "La plantilla predeterminada se ha guardado correctamente.",
                parent=self.root
            )
        return True

    def _load_default_into_provider_editor(self):
        default = self._email_templates["default"]
        self.subject_entry.delete(0, tk.END)
        self.subject_entry.insert(0, default.get("subject", ""))
        self.body_editor.set_html(default.get("body", ""))

    def _load_provider_email(self, account):
        # Las personalizaciones por proveedor solo viven durante esta sesión.
        # En una nueva apertura, todos parten de la plantilla predeterminada.
        provider_data = self._provider_email_session.get(account)
        if provider_data is None:
            self._load_default_into_provider_editor()
            return

        self.subject_entry.delete(0, tk.END)
        self.subject_entry.insert(0, provider_data.get("subject", ""))
        self.body_editor.set_html(provider_data.get("body", ""))

    def _save_current_provider_email(self):
        if not self._selected_provider:
            return True

        # Guardado automático en memoria al cambiar de proveedor.
        # No se escribe en email_templates.json.
        self._provider_email_session[self._selected_provider] = {
            "subject": self.subject_entry.get(),
            "body": self.body_editor.get_html()
        }
        return True

    def _get_email_html_with_signature(self, body_html):
        """Devuelve el cuerpo HTML seguido de la firma importada."""
        signature_html = (self._signature_html or "").strip()
        body_html = (body_html or "").strip()
        if not signature_html:
            return body_html
        if not body_html:
            return signature_html
        return body_html + "<br><br>" + signature_html

    def _on_close(self):
        # Solo persiste la plantilla predeterminada.
        # La firma se guarda explícitamente con su botón; los mensajes
        # personalizados de proveedores se descartan al cerrar.
        if self._save_current_provider_email():
            if self._save_default_email(show_confirmation=False):
                self.root.destroy()

    def _toggle_all_lines(self):
        self._set_all_lines_selected(self.all_lines_var.get())

    def _set_all_lines_selected(self, selected):
        if not self._selected_provider:
            return

        rows = self._lines_by_provider.get(self._selected_provider, [])
        self._line_selected = {str(index): bool(selected) for index, _row in enumerate(rows)}
        self._line_selection_by_provider[self._selected_provider] = dict(self._line_selected)
        self.all_lines_var.set(bool(selected))
        self._refresh_lines()

    def load_initial_data(self):

        try:
            contacts = self.repository.get_contacts()
            suppliers = self.repository.get_suppliers()

            self._contacts_by_account = {
                str(row.get("account")): row
                for row in contacts
            }

            self.set_providers(suppliers)

        except Exception as exc:

            messagebox.showerror(
                "Error de carga",
                str(exc),
                parent=self.root
            )

    def set_providers(
        self,
        providers: Iterable[Mapping[str, Any]]
    ):

        self._providers = [
            dict(row)
            for row in providers
        ]

        for row in self._providers:

            account = str(row["account"])

            self._provider_draft_enabled.setdefault(
                account,
                False
            )

        self._refresh_providers()

    def set_provider_lines(
        self,
        account: str,
        lines: Iterable[Mapping[str, Any]]
    ):

        self._lines_by_provider[str(account)] = [
            dict(row)
            for row in lines
        ]

        if self._selected_provider == str(account):
            self._refresh_lines()

    def _refresh_providers(self):

        selected = self._selected_provider

        self.providers_tree.delete(
            *self.providers_tree.get_children()
        )

        query = self.search_var.get().strip().casefold()

        for row in self._sorted_providers():

            account = str(row["account"])
            name = str(row["name"])

            if (
                query
                and query not in account.casefold()
                and query not in name.casefold()
            ):
                continue

            included = self._provider_draft_enabled.get(
                account,
                False
            )

            self.providers_tree.insert(
                "",
                tk.END,
                iid=account,
                values=(
                    "☑" if included else "☐",
                    account,
                    name,
                    row.get("orders", "")
                ),
                tags=("selected_bold",) if included else ()
            )

        if selected and self.providers_tree.exists(selected):

            self.providers_tree.selection_set(selected)
            self.providers_tree.focus(selected)

    def _sorted_providers(self):

        key = {
            "account": "account",
            "name": "name",
            "orders": "orders"
        }.get(
            self._provider_sort_column,
            "name"
        )

        return sorted(
            self._providers,
            key=lambda row: str(
                row.get(key, "")
            ).casefold(),
            reverse=self._provider_sort_reverse
        )

    def _sort_providers(self, column):

        if column == "include":
            return

        if self._provider_sort_column == column:

            self._provider_sort_reverse = (
                not self._provider_sort_reverse
            )

        else:

            self._provider_sort_column = column
            self._provider_sort_reverse = False

        self._refresh_providers()

    def _on_search_changed(self, *_args):

        self._refresh_providers()

    def _on_provider_click(self, event):

        row_id = self.providers_tree.identify_row(event.y)

        if (
            row_id
            and self.providers_tree.identify_column(event.x) == "#1"
        ):

            self._provider_draft_enabled[row_id] = (
                not self._provider_draft_enabled.get(
                    row_id,
                    False
                )
            )

            self._refresh_providers()

            return "break"

        return None

    def _on_provider_selected(self, _event):

        selection = self.providers_tree.selection()

        if not selection:

            if self._selected_provider:
                self._line_selection_by_provider[self._selected_provider] = dict(
                    self._line_selected
                )
                self._line_column_selection_by_provider[self._selected_provider] = dict(
                    self._line_column_selected
                )
            if not self._save_current_provider_email():
                return
            self._selected_provider = None
            self._refresh_lines()

            return

        account = selection[0]
        previous_account = self._selected_provider
        if previous_account and previous_account != account:
            self._line_selection_by_provider[previous_account] = dict(self._line_selected)
            self._line_column_selection_by_provider[previous_account] = dict(
                self._line_column_selected
            )
            if not self._save_current_provider_email():
                return

        self._selected_provider = account
        self._line_column_selected = dict(
            self._line_column_selection_by_provider.get(
                account,
                {col: True for col in self._line_columns}
            )
        )
        for filter_var in self._line_filter_vars.values():
            filter_var.set("")
        self._load_provider_email(account)

        self.email_provider_label.configure(
            text=f"Proveedor seleccionado: {account}"
        )

        if self.repository is not None:

            try:

                lines = self.repository.get_order_lines(account)

                self.set_provider_lines(
                    account,
                    lines
                )

            except Exception as exc:

                self.set_provider_lines(account, [])

                messagebox.showerror(
                    "Error al cargar líneas",
                    str(exc),
                    parent=self.root
                )

                return

        self._refresh_lines()

    def _refresh_lines(self):
        self._visible_line_rows = []
        self._line_selected.clear()

        if not self._selected_provider:
            self.lines_status.configure(
                text="Selecciona un proveedor para ver sus líneas."
            )
            self.all_lines_var.set(False)
            self._draw_lines_table()
            return

        rows = self._lines_by_provider.get(self._selected_provider, [])
        saved_selection = self._line_selection_by_provider.get(self._selected_provider, {})
        filters = {
            col: self._line_filter_vars[col].get().strip()
            for col in self._line_columns
        }

        indexed_rows = list(enumerate(rows))
        if self._line_sort_column:
            indexed_rows.sort(
                key=lambda item: self._line_sort_key(
                    item[1].get(self._line_sort_column, ""),
                    self._line_sort_column
                ),
                reverse=self._line_sort_reverse
            )

        for index, row in indexed_rows:
            row_id = str(index)
            is_selected = saved_selection.get(row_id, True)
            self._line_selected[row_id] = is_selected

            matches = True
            for col, expression in filters.items():
                if not expression:
                    continue
                value = row.get(col, "")
                if col == "quantity":
                    matches = self._match_numeric_filter(value, expression)
                elif col in ("date", "order_date", "delivery_date"):
                    matches = self._match_date_filter(value, expression)
                else:
                    matches = self._match_text_filter(value, expression)
                if not matches:
                    break
            if matches:
                self._visible_line_rows.append((row_id, row))

        visible_count = len(self._visible_line_rows)
        self.all_lines_var.set(
            bool(visible_count) and all(
                self._line_selected.get(row_id, False)
                for row_id, _row in self._visible_line_rows
            )
        )
        self.lines_status.configure(
            text=(
                f"{visible_count} de {len(rows)} líneas visibles del proveedor "
                f"{self._selected_provider}."
                if rows else "No hay líneas cargadas para este proveedor."
            )
        )
        self._draw_lines_table()

    def _on_line_click(self, event):
        return None

    def _on_line_navigation(self, event):
        """Mueve la fila activa con ↑/↓ y permite saltar al inicio/final."""
        if not self._visible_line_rows:
            return "break"

        visible_ids = [row_id for row_id, _row in self._visible_line_rows]
        if self._last_focused_line_id in visible_ids:
            current_index = visible_ids.index(self._last_focused_line_id)
        else:
            current_index = 0

        if event.keysym == "Up":
            next_index = max(0, current_index - 1)
        elif event.keysym == "Down":
            next_index = min(len(visible_ids) - 1, current_index + 1)
        elif event.keysym == "Home":
            next_index = 0
        elif event.keysym == "End":
            next_index = len(visible_ids) - 1
        else:
            return "break"

        self._last_focused_line_id = visible_ids[next_index]
        row_top = next_index * 28
        row_bottom = row_top + 28
        visible_top = self.lines_canvas.canvasy(0)
        visible_bottom = visible_top + self.lines_canvas.winfo_height()

        if row_top < visible_top:
            self.lines_canvas.yview_moveto(row_top / max(1, len(visible_ids) * 28))
            self.include_canvas.yview_moveto(row_top / max(1, len(visible_ids) * 28))
        elif row_bottom > visible_bottom:
            target_top = max(0, row_bottom - self.lines_canvas.winfo_height())
            self.lines_canvas.yview_moveto(target_top / max(1, len(visible_ids) * 28))
            self.include_canvas.yview_moveto(target_top / max(1, len(visible_ids) * 28))

        self._draw_lines_table()
        return "break"

    def _on_line_enter(self, _event):
        """Alterna la inclusión de la fila activa con Enter."""
        row_id = self._last_focused_line_id
        if row_id and row_id in self._line_selected:
            self._toggle_line_selection(row_id)
        return "break"



