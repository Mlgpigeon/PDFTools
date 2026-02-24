#!/usr/bin/env python3
"""
PDF Toolkit — Herramienta completa de manipulación de PDFs.

Funciones: Unir, Dividir, Extraer, Rotar, Imágenes→PDF, DOCX→PDF,
           Comprimir, Proteger/Desproteger.

Ejecutar:  python pdf_toolkit.py
CLI:       python pdf_toolkit.py --merge f1.pdf f2.pdf -o salida.pdf
"""

import os
import sys
import json
import re
import threading
import queue
import subprocess
import shutil
import argparse
import tempfile
from pathlib import Path
from typing import Callable, Any

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog

# ═══════════════════════════════════════════════════════════════════════════════
# DETECCIÓN DE DEPENDENCIAS
# ═══════════════════════════════════════════════════════════════════════════════

DND_AVAILABLE = True
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
except Exception:
    DND_AVAILABLE = False
    TkinterDnD = None
    DND_FILES = None

PYPDF_AVAILABLE = True
try:
    from pypdf import PdfReader, PdfWriter
except Exception:
    PYPDF_AVAILABLE = False
    PdfReader = None
    PdfWriter = None

PILLOW_AVAILABLE = True
try:
    from PIL import Image
except Exception:
    PILLOW_AVAILABLE = False
    Image = None

DOCX2PDF_AVAILABLE = True
try:
    from docx2pdf import convert as docx2pdf_convert
except Exception:
    DOCX2PDF_AVAILABLE = False
    docx2pdf_convert = None

LIBREOFFICE_PATH = shutil.which("soffice") or shutil.which("libreoffice")

GHOSTSCRIPT_PATH = None
for gs_name in ("gs", "gswin64c", "gswin32c", "gswin64c.exe", "gswin32c.exe"):
    found = shutil.which(gs_name)
    if found:
        GHOSTSCRIPT_PATH = found
        break

# ═══════════════════════════════════════════════════════════════════════════════
# CONFIGURACIÓN PERSISTENTE
# ═══════════════════════════════════════════════════════════════════════════════

CONFIG_PATH = os.path.join(os.path.expanduser("~"), ".pdf_toolkit_config.json")

def load_config() -> dict:
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}

def save_config(cfg: dict):
    try:
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=2, ensure_ascii=False)
    except Exception:
        pass

# ═══════════════════════════════════════════════════════════════════════════════
# UTILIDADES
# ═══════════════════════════════════════════════════════════════════════════════

def collect_files(paths: list[str], extensions: list[str]) -> list[str]:
    """Recopila archivos con las extensiones dadas desde rutas (archivos o carpetas)."""
    out: list[str] = []
    ext_set = {e.lower().lstrip(".") for e in extensions}
    for p in paths:
        p = p.strip().strip('"')
        if not p:
            continue
        if os.path.isdir(p):
            for root, _, files in os.walk(p):
                for f in files:
                    fp = os.path.join(root, f)
                    if os.path.splitext(fp)[1].lower().lstrip(".") in ext_set:
                        out.append(os.path.abspath(fp))
        elif os.path.isfile(p):
            if os.path.splitext(p)[1].lower().lstrip(".") in ext_set:
                out.append(os.path.abspath(p))
    return out


def parse_page_ranges(text: str, max_page: int) -> list[int]:
    """
    Parsea rangos como '1-3, 5, 8-12' y devuelve lista de índices 0-based.
    Lanza ValueError si hay páginas fuera de rango.
    """
    pages: list[int] = []
    text = text.strip()
    if not text:
        return pages

    for part in text.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            sides = part.split("-", 1)
            start = int(sides[0].strip())
            end = int(sides[1].strip())
            if start < 1 or end > max_page or start > end:
                raise ValueError(f"Rango inválido: {part} (el PDF tiene {max_page} páginas)")
            pages.extend(range(start - 1, end))
        else:
            num = int(part.strip())
            if num < 1 or num > max_page:
                raise ValueError(f"Página {num} fuera de rango (el PDF tiene {max_page} páginas)")
            pages.append(num - 1)
    return pages


def setup_dnd(widget, root, callback):
    """Configura drag & drop en un widget si está disponible."""
    if DND_AVAILABLE:
        try:
            widget.drop_target_register(DND_FILES)
            widget.dnd_bind("<<Drop>>", callback)
            return True
        except Exception:
            pass
    return False


def get_pdf_page_count(filepath: str) -> int:
    """Obtiene el número de páginas de un PDF."""
    if not PYPDF_AVAILABLE:
        return 0
    try:
        reader = PdfReader(filepath)
        return len(reader.pages)
    except Exception:
        return 0


def format_size(size_bytes: int) -> str:
    """Formatea bytes a cadena legible."""
    if size_bytes < 1024:
        return f"{size_bytes} B"
    elif size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.1f} KB"
    else:
        return f"{size_bytes / (1024 * 1024):.1f} MB"


def ask_pdf_password(parent, filepath: str) -> str | None:
    """Pide contraseña para un PDF protegido."""
    nombre = os.path.basename(filepath)
    pwd = simpledialog.askstring(
        "PDF Protegido",
        f"El archivo '{nombre}' está protegido.\nIntroduce la contraseña:",
        parent=parent,
        show="*"
    )
    return pwd


# ═══════════════════════════════════════════════════════════════════════════════
# BACKENDS (funciones puras, sin UI)
# ═══════════════════════════════════════════════════════════════════════════════

def merge_pdfs(files: list[str], output_path: str, progress_cb=None, passwords: dict = None):
    """Une varios PDFs en uno solo."""
    writer = PdfWriter()
    for i, f in enumerate(files):
        reader = PdfReader(f)
        if reader.is_encrypted:
            pwd = (passwords or {}).get(f, "")
            reader.decrypt(pwd)
        for page in reader.pages:
            writer.add_page(page)
        if progress_cb:
            progress_cb(i + 1, len(files))
    with open(output_path, "wb") as out:
        writer.write(out)


def split_pdf(filepath: str, mode: str, output_dir: str, prefix: str = "parte",
              n_pages: int = 1, ranges_text: str = "", password: str = "",
              progress_cb=None):
    """
    Divide un PDF.
    mode: 'each_n' | 'page_by_page' | 'ranges'
    """
    reader = PdfReader(filepath)
    if reader.is_encrypted:
        reader.decrypt(password)
    total = len(reader.pages)

    os.makedirs(output_dir, exist_ok=True)
    files_created = []

    if mode == "page_by_page":
        for i in range(total):
            writer = PdfWriter()
            writer.add_page(reader.pages[i])
            out_name = os.path.join(output_dir, f"{prefix}_{i + 1:03d}.pdf")
            with open(out_name, "wb") as out:
                writer.write(out)
            files_created.append(out_name)
            if progress_cb:
                progress_cb(i + 1, total)

    elif mode == "each_n":
        part = 1
        for start in range(0, total, n_pages):
            writer = PdfWriter()
            end = min(start + n_pages, total)
            for i in range(start, end):
                writer.add_page(reader.pages[i])
            out_name = os.path.join(output_dir, f"{prefix}_{part:03d}.pdf")
            with open(out_name, "wb") as out:
                writer.write(out)
            files_created.append(out_name)
            if progress_cb:
                progress_cb(end, total)
            part += 1

    elif mode == "ranges":
        page_groups = [r.strip() for r in ranges_text.split(",") if r.strip()]
        group_idx = 1
        done = 0
        for group in page_groups:
            writer = PdfWriter()
            if "-" in group:
                sides = group.split("-", 1)
                s, e = int(sides[0].strip()) - 1, int(sides[1].strip())
                for i in range(s, min(e, total)):
                    writer.add_page(reader.pages[i])
            else:
                idx = int(group.strip()) - 1
                if 0 <= idx < total:
                    writer.add_page(reader.pages[idx])
            out_name = os.path.join(output_dir, f"{prefix}_{group_idx:03d}.pdf")
            with open(out_name, "wb") as out:
                writer.write(out)
            files_created.append(out_name)
            done += 1
            if progress_cb:
                progress_cb(done, len(page_groups))
            group_idx += 1

    return files_created


def extract_pages(filepath: str, ranges_text: str, output_path: str,
                  password: str = "", progress_cb=None):
    """Extrae páginas de un PDF a un nuevo archivo."""
    reader = PdfReader(filepath)
    if reader.is_encrypted:
        reader.decrypt(password)
    total = len(reader.pages)
    indices = parse_page_ranges(ranges_text, total)

    writer = PdfWriter()
    for i, idx in enumerate(indices):
        writer.add_page(reader.pages[idx])
        if progress_cb:
            progress_cb(i + 1, len(indices))

    with open(output_path, "wb") as out:
        writer.write(out)


def rotate_pages(filepath: str, page_selection: str, degrees: int,
                 output_path: str, ranges_text: str = "", password: str = "",
                 progress_cb=None):
    """
    Rota páginas de un PDF.
    page_selection: 'all' | 'even' | 'odd' | 'custom'
    """
    reader = PdfReader(filepath)
    if reader.is_encrypted:
        reader.decrypt(password)
    total = len(reader.pages)

    writer = PdfWriter()
    writer.clone_document_from_reader(reader)

    if page_selection == "all":
        indices = list(range(total))
    elif page_selection == "even":
        indices = list(range(1, total, 2))
    elif page_selection == "odd":
        indices = list(range(0, total, 2))
    elif page_selection == "custom":
        indices = parse_page_ranges(ranges_text, total)
    else:
        indices = list(range(total))

    for i, idx in enumerate(indices):
        writer.pages[idx].rotate(degrees)
        if progress_cb:
            progress_cb(i + 1, len(indices))

    with open(output_path, "wb") as out:
        writer.write(out)


def images_to_pdf(image_files: list[str], output_path: str,
                  page_size: str = "a4", orientation: str = "auto",
                  margin_mm: float = 10, progress_cb=None):
    """Convierte una lista de imágenes a un PDF."""
    # Tamaños en puntos (72 dpi)
    SIZES = {
        "a4": (595.28, 841.89),
        "letter": (612, 792),
        "fit": None,
    }

    target = SIZES.get(page_size.lower(), SIZES["a4"])
    margin_pt = margin_mm * 72 / 25.4

    processed: list[Image.Image] = []

    for i, img_path in enumerate(image_files):
        img = Image.open(img_path)

        # CMYK → RGB
        if img.mode == "CMYK":
            img = img.convert("RGB")
        elif img.mode != "RGB":
            img = img.convert("RGB")

        if target is None:
            # Ajustar al tamaño de la imagen
            processed.append(img)
        else:
            pw, ph = target

            # Orientación
            if orientation == "landscape":
                pw, ph = max(pw, ph), min(pw, ph)
            elif orientation == "portrait":
                pw, ph = min(pw, ph), max(pw, ph)
            elif orientation == "auto":
                if img.width > img.height:
                    pw, ph = max(pw, ph), min(pw, ph)

            # Crear canvas y centrar imagen
            usable_w = pw - 2 * margin_pt
            usable_h = ph - 2 * margin_pt
            scale = min(usable_w / img.width, usable_h / img.height)
            new_w = int(img.width * scale)
            new_h = int(img.height * scale)
            img_resized = img.resize((new_w, new_h), Image.LANCZOS)

            canvas = Image.new("RGB", (int(pw), int(ph)), (255, 255, 255))
            x = int((pw - new_w) / 2)
            y = int((ph - new_h) / 2)
            canvas.paste(img_resized, (x, y))
            processed.append(canvas)

        if progress_cb:
            progress_cb(i + 1, len(image_files))

    if processed:
        processed[0].save(output_path, "PDF", save_all=True,
                          append_images=processed[1:], resolution=150)


def convert_with_word(docx_path: str, outdir: str | None) -> None:
    """Usa docx2pdf → Microsoft Word COM (Windows/macOS)."""
    if sys.platform.startswith("win"):
        try:
            import pythoncom
            pythoncom.CoInitialize()
        except Exception:
            pythoncom = None
    else:
        pythoncom = None

    try:
        if outdir:
            docx2pdf_convert(docx_path, outdir)
        else:
            docx2pdf_convert(docx_path)
    finally:
        if sys.platform.startswith("win") and pythoncom is not None:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass


def convert_with_libreoffice(docx_path: str, outdir: str | None) -> None:
    """Usa LibreOffice headless: soffice --headless --convert-to pdf."""
    target_dir = outdir or os.path.dirname(docx_path)
    os.makedirs(target_dir, exist_ok=True)

    soffice = LIBREOFFICE_PATH
    if not soffice:
        raise RuntimeError(
            "No encuentro 'soffice' (LibreOffice) en el PATH.\n"
            "Instala LibreOffice o añade soffice.exe al PATH."
        )

    cmd = [
        soffice, "--headless", "--nologo", "--nolockcheck", "--norestore",
        "--convert-to", "pdf", "--outdir", target_dir, docx_path,
    ]
    p = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    if p.returncode != 0:
        raise RuntimeError(p.stderr.strip() or f"LibreOffice falló (code {p.returncode}).")


def compress_pdf(filepath: str, output_path: str, quality: str = "ebook",
                 progress_cb=None):
    """Comprime un PDF usando Ghostscript."""
    if not GHOSTSCRIPT_PATH:
        raise RuntimeError("Ghostscript no encontrado. Instálalo para usar esta función.")

    cmd = [
        GHOSTSCRIPT_PATH,
        "-sDEVICE=pdfwrite",
        "-dCompatibilityLevel=1.4",
        f"-dPDFSETTINGS=/{quality}",
        "-dNOPAUSE", "-dBATCH", "-dQUIET",
        f"-sOutputFile={output_path}",
        filepath,
    ]
    p = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    if p.returncode != 0:
        raise RuntimeError(p.stderr.strip() or f"Ghostscript falló (code {p.returncode}).")
    if progress_cb:
        progress_cb(1, 1)


def compress_pdf_pypdf(filepath: str, output_path: str, password: str = "",
                       progress_cb=None):
    """Compresión básica con pypdf (sin Ghostscript)."""
    reader = PdfReader(filepath)
    if reader.is_encrypted:
        reader.decrypt(password)
    writer = PdfWriter()
    total = len(reader.pages)
    for i, page in enumerate(reader.pages):
        page.compress_content_streams()
        writer.add_page(page)
        if progress_cb:
            progress_cb(i + 1, total)
    with open(output_path, "wb") as out:
        writer.write(out)


def protect_pdf(filepath: str, output_path: str, user_password: str,
                owner_password: str = "", permissions: dict = None,
                password: str = "", progress_cb=None):
    """Protege un PDF con contraseña."""
    reader = PdfReader(filepath)
    if reader.is_encrypted:
        reader.decrypt(password)
    writer = PdfWriter()
    writer.clone_document_from_reader(reader)

    from pypdf.constants import UserAccessPermissions
    perms = UserAccessPermissions.all()
    if permissions:
        if not permissions.get("print", True):
            perms &= ~UserAccessPermissions.PRINT
            perms &= ~UserAccessPermissions.PRINT_TO_REPRESENTATION
        if not permissions.get("copy", True):
            perms &= ~UserAccessPermissions.EXTRACT_TEXT_AND_GRAPHICS
        if not permissions.get("modify", True):
            perms &= ~UserAccessPermissions.MODIFY

    writer.encrypt(
        user_password=user_password,
        owner_password=owner_password or user_password,
        algorithm="AES-256",
        permissions=perms,
    )
    with open(output_path, "wb") as out:
        writer.write(out)
    if progress_cb:
        progress_cb(1, 1)


def unprotect_pdf(filepath: str, output_path: str, password: str,
                  progress_cb=None):
    """Quita la protección de un PDF."""
    reader = PdfReader(filepath)
    if not reader.decrypt(password):
        raise ValueError("Contraseña incorrecta.")
    writer = PdfWriter()
    writer.clone_document_from_reader(reader)
    with open(output_path, "wb") as out:
        writer.write(out)
    if progress_cb:
        progress_cb(1, 1)


def add_watermark_text(filepath: str, output_path: str, text: str,
                       font_size: int = 50, opacity: float = 0.3,
                       rotation: int = 45, password: str = "",
                       progress_cb=None):
    """Añade una marca de agua de texto a cada página."""
    from pypdf import Transformation
    import io

    # Crear PDF con marca de agua usando reportlab si está disponible,
    # si no, usar un enfoque más simple con pypdf
    reader = PdfReader(filepath)
    if reader.is_encrypted:
        reader.decrypt(password)
    writer = PdfWriter()

    # Intentar usar reportlab para la marca de agua
    try:
        from reportlab.pdfgen import canvas as rl_canvas
        from reportlab.lib.pagesizes import A4

        total = len(reader.pages)
        for i, page in enumerate(reader.pages):
            # Crear marca de agua del tamaño de la página
            page_w = float(page.mediabox.width)
            page_h = float(page.mediabox.height)

            packet = io.BytesIO()
            c = rl_canvas.Canvas(packet, pagesize=(page_w, page_h))
            c.setFont("Helvetica", font_size)
            c.setFillAlpha(opacity)
            c.setFillColorRGB(0.5, 0.5, 0.5)
            c.saveState()
            c.translate(page_w / 2, page_h / 2)
            c.rotate(rotation)
            c.drawCentredString(0, 0, text)
            c.restoreState()
            c.save()
            packet.seek(0)

            wm_reader = PdfReader(packet)
            page.merge_page(wm_reader.pages[0])
            writer.add_page(page)

            if progress_cb:
                progress_cb(i + 1, total)

    except ImportError:
        # Sin reportlab, simplemente clonar (no se puede añadir texto fácilmente)
        raise RuntimeError(
            "Se necesita 'reportlab' para marcas de agua de texto.\n"
            "Instala con: pip install reportlab"
        )

    with open(output_path, "wb") as out:
        writer.write(out)


def set_pdf_metadata(filepath: str, output_path: str, metadata: dict,
                     password: str = "", progress_cb=None):
    """Establece metadatos (título, autor, etc.) en un PDF."""
    reader = PdfReader(filepath)
    if reader.is_encrypted:
        reader.decrypt(password)
    writer = PdfWriter()
    writer.clone_document_from_reader(reader)
    writer.add_metadata(metadata)
    with open(output_path, "wb") as out:
        writer.write(out)
    if progress_cb:
        progress_cb(1, 1)


# ═══════════════════════════════════════════════════════════════════════════════
# TASK RUNNER (threading genérico)
# ═══════════════════════════════════════════════════════════════════════════════

class TaskRunner:
    """Ejecuta una tarea en un hilo separado, reporta progreso vía queue."""

    def __init__(self, task_fn: Callable, msg_queue: queue.Queue,
                 stop_flag: threading.Event):
        self.task_fn = task_fn
        self.q = msg_queue
        self.stop_flag = stop_flag
        self.thread: threading.Thread | None = None

    def start(self, *args, **kwargs):
        self.stop_flag.clear()
        self.thread = threading.Thread(
            target=self._run, args=args, kwargs=kwargs, daemon=True
        )
        self.thread.start()

    def _run(self, *args, **kwargs):
        try:
            result = self.task_fn(*args, **kwargs)
            self.q.put(("done_ok", result))
        except Exception as e:
            self.q.put(("done_error", str(e)))

    def is_alive(self) -> bool:
        return self.thread is not None and self.thread.is_alive()


# ═══════════════════════════════════════════════════════════════════════════════
# UI — WIDGET BASE REUTILIZABLE
# ═══════════════════════════════════════════════════════════════════════════════

class BaseTab(ttk.Frame):
    """Pestaña base con listbox + dnd + progreso + estado."""

    def __init__(self, parent, app, accepted_extensions: list[str],
                 multi_file: bool = True, **kwargs):
        super().__init__(parent, **kwargs)
        self.app = app
        self.accepted_extensions = accepted_extensions
        self.multi_file = multi_file

        self.files: list[str] = []
        self.file_set: set[str] = set()
        self.q: queue.Queue = queue.Queue()
        self.stop_flag = threading.Event()
        self.runner: TaskRunner | None = None

        self._build_base_ui()
        self.build_options(self.options_frame)
        self._tick()

    def _build_base_ui(self):
        main = ttk.Frame(self, padding=10)
        main.pack(fill="both", expand=True)

        # --- Zona de entrada ---
        if self.multi_file:
            ext_str = ", ".join(f".{e}" for e in self.accepted_extensions)
            input_frame = ttk.LabelFrame(
                main,
                text=f"Arrastra archivos aquí ({ext_str})",
                padding=8
            )
            input_frame.pack(fill="both", expand=True)

            # Listbox + scrollbar
            list_container = ttk.Frame(input_frame)
            list_container.pack(fill="both", expand=True)

            self.listbox = tk.Listbox(list_container, selectmode=tk.EXTENDED,
                                      font=("Consolas", 9))
            self.listbox.pack(side="left", fill="both", expand=True)

            sb = ttk.Scrollbar(list_container, orient="vertical",
                               command=self.listbox.yview)
            sb.pack(side="right", fill="y")
            self.listbox.configure(yscrollcommand=sb.set)

            setup_dnd(self.listbox, self.app.root, self._on_drop)

            # Botones de lista
            btn_frame = ttk.Frame(input_frame)
            btn_frame.pack(fill="x", pady=(6, 0))

            ttk.Button(btn_frame, text="Añadir…",
                       command=self._add_files_dialog).pack(side="left")
            ttk.Button(btn_frame, text="Quitar",
                       command=self._remove_selected).pack(side="left", padx=4)
            ttk.Button(btn_frame, text="Limpiar",
                       command=self._clear).pack(side="left")

            # Botones de reorden
            ttk.Separator(btn_frame, orient="vertical").pack(
                side="left", fill="y", padx=8)
            ttk.Button(btn_frame, text="▲", width=3,
                       command=self._move_up).pack(side="left")
            ttk.Button(btn_frame, text="▼", width=3,
                       command=self._move_down).pack(side="left", padx=(4, 0))

        else:
            # Archivo único
            input_frame = ttk.LabelFrame(main, text="Archivo de entrada", padding=8)
            input_frame.pack(fill="x")

            file_row = ttk.Frame(input_frame)
            file_row.pack(fill="x")

            self.file_entry = ttk.Entry(file_row, font=("Consolas", 9))
            self.file_entry.pack(side="left", fill="x", expand=True)

            ttk.Button(file_row, text="Seleccionar…",
                       command=self._select_single_file).pack(side="left", padx=(6, 0))

            setup_dnd(self.file_entry, self.app.root, self._on_drop_single)

            self.listbox = None
            self.page_info_label = ttk.Label(input_frame, text="",
                                             foreground="gray")
            self.page_info_label.pack(anchor="w", pady=(4, 0))

        # --- Opciones (a implementar por cada pestaña) ---
        self.options_frame = ttk.LabelFrame(main, text="Opciones", padding=8)
        self.options_frame.pack(fill="x", pady=(8, 0))

        # --- Zona de salida ---
        out_frame = ttk.LabelFrame(main, text="Salida", padding=8)
        out_frame.pack(fill="x", pady=(8, 0))

        self.out_row = ttk.Frame(out_frame)
        self.out_row.pack(fill="x")

        self.out_var = tk.StringVar(value="")
        self.out_entry = ttk.Entry(self.out_row, textvariable=self.out_var,
                                   font=("Consolas", 9))
        self.out_entry.pack(side="left", fill="x", expand=True)

        self.btn_out_browse = ttk.Button(self.out_row, text="Elegir…",
                                         command=self._choose_output)
        self.btn_out_browse.pack(side="left", padx=(6, 0))

        self.output_is_folder = True  # Subclases pueden cambiar esto

        # --- Progreso + estado ---
        bottom = ttk.Frame(main, padding=(0, 8, 0, 0))
        bottom.pack(fill="x")

        self.progress = ttk.Progressbar(bottom, mode="determinate")
        self.progress.pack(fill="x")

        status_row = ttk.Frame(bottom)
        status_row.pack(fill="x", pady=(6, 0))

        self.status_label = ttk.Label(status_row, text="Listo.",
                                      foreground="gray")
        self.status_label.pack(side="left", fill="x", expand=True)

        self.size_label = ttk.Label(status_row, text="", foreground="gray")
        self.size_label.pack(side="right")

        # --- Botones ejecutar/parar ---
        action_row = ttk.Frame(bottom)
        action_row.pack(fill="x", pady=(6, 0))

        self.btn_run = ttk.Button(action_row, text="Ejecutar",
                                  command=self._on_run)
        self.btn_run.pack(side="left")

        self.btn_stop = ttk.Button(action_row, text="Parar",
                                   command=self._on_stop, state="disabled")
        self.btn_stop.pack(side="left", padx=6)

    # --- Drag & Drop ---
    def _on_drop(self, event):
        try:
            paths = list(self.app.root.tk.splitlist(event.data))
        except Exception:
            paths = [event.data]
        self._add_paths(paths)

    def _on_drop_single(self, event):
        try:
            paths = list(self.app.root.tk.splitlist(event.data))
        except Exception:
            paths = [event.data]
        files = collect_files(paths, self.accepted_extensions)
        if files:
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, files[0])
            self.files = [files[0]]
            self._update_page_info()

    # --- File management ---
    def _add_files_dialog(self):
        ext_patterns = " ".join(f"*.{e}" for e in self.accepted_extensions)
        files = filedialog.askopenfilenames(
            title="Selecciona archivos",
            filetypes=[("Archivos soportados", ext_patterns), ("Todos", "*.*")]
        )
        if files:
            self._add_paths(list(files))

    def _select_single_file(self):
        ext_patterns = " ".join(f"*.{e}" for e in self.accepted_extensions)
        f = filedialog.askopenfilename(
            title="Selecciona un archivo",
            filetypes=[("Archivos soportados", ext_patterns), ("Todos", "*.*")]
        )
        if f:
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, f)
            self.files = [f]
            self._update_page_info()

    def _update_page_info(self):
        """Muestra información de páginas para un archivo PDF."""
        if not self.multi_file and hasattr(self, "page_info_label") and self.files:
            filepath = self.files[0]
            if filepath.lower().endswith(".pdf"):
                count = get_pdf_page_count(filepath)
                size = format_size(os.path.getsize(filepath)) if os.path.exists(filepath) else ""
                self.page_info_label.configure(
                    text=f"{count} páginas — {size}")
            else:
                self.page_info_label.configure(text="")

    def _add_paths(self, paths: list[str]):
        docs = collect_files(paths, self.accepted_extensions)
        added = 0
        for p in docs:
            if p not in self.file_set:
                self.file_set.add(p)
                self.files.append(p)
                display = self._format_listbox_entry(p)
                self.listbox.insert(tk.END, display)
                added += 1

        if added:
            self.set_status(f"Añadidos: {added}. Total: {len(self.files)}")
        else:
            self.set_status("No se encontraron archivos nuevos con las extensiones esperadas.")

    def _format_listbox_entry(self, filepath: str) -> str:
        """Formatea la entrada del listbox mostrando páginas y tamaño."""
        name = os.path.basename(filepath)
        info_parts = []

        if filepath.lower().endswith(".pdf"):
            count = get_pdf_page_count(filepath)
            if count:
                info_parts.append(f"{count}pp")

        try:
            size = format_size(os.path.getsize(filepath))
            info_parts.append(size)
        except Exception:
            pass

        if info_parts:
            return f"{name}  [{', '.join(info_parts)}]"
        return name

    def _remove_selected(self):
        if not self.listbox:
            return
        sel = list(self.listbox.curselection())
        if not sel:
            return
        for idx in reversed(sel):
            p = self.files[idx]
            self.listbox.delete(idx)
            self.file_set.discard(p)
            self.files.pop(idx)
        self.set_status(f"Total: {len(self.files)}")

    def _clear(self):
        if not self.listbox:
            return
        self.listbox.delete(0, tk.END)
        self.files.clear()
        self.file_set.clear()
        self.set_status("Lista vacía.")

    def _move_up(self):
        if not self.listbox:
            return
        sel = list(self.listbox.curselection())
        if not sel or sel[0] == 0:
            return
        for idx in sel:
            self.files[idx - 1], self.files[idx] = self.files[idx], self.files[idx - 1]
        self._refresh_listbox()
        for idx in sel:
            self.listbox.selection_set(idx - 1)

    def _move_down(self):
        if not self.listbox:
            return
        sel = list(self.listbox.curselection())
        if not sel or sel[-1] >= len(self.files) - 1:
            return
        for idx in reversed(sel):
            self.files[idx + 1], self.files[idx] = self.files[idx], self.files[idx + 1]
        self._refresh_listbox()
        for idx in sel:
            self.listbox.selection_set(idx + 1)

    def _refresh_listbox(self):
        if not self.listbox:
            return
        self.listbox.delete(0, tk.END)
        for f in self.files:
            self.listbox.insert(tk.END, self._format_listbox_entry(f))

    # --- Output ---
    def _choose_output(self):
        if self.output_is_folder:
            d = filedialog.askdirectory(title="Carpeta de salida")
            if d:
                self.out_var.set(os.path.abspath(d))
        else:
            f = filedialog.asksaveasfilename(
                title="Guardar como",
                defaultextension=".pdf",
                filetypes=[("PDF", "*.pdf"), ("Todos", "*.*")]
            )
            if f:
                self.out_var.set(os.path.abspath(f))

    # --- Status ---
    def set_status(self, text: str, color: str = "gray"):
        self.status_label.configure(text=text, foreground=color)

    def set_size_info(self, text: str):
        self.size_label.configure(text=text)

    # --- Execution ---
    def _on_run(self):
        if self.runner and self.runner.is_alive():
            return
        self.execute()

    def _on_stop(self):
        self.stop_flag.set()
        self.set_status("Parando…", "orange")

    def run_task(self, fn: Callable, *args, **kwargs):
        """Lanza una función en un hilo separado."""
        self.progress["value"] = 0
        self.btn_run.configure(state="disabled")
        self.btn_stop.configure(state="normal")
        self.stop_flag.clear()

        self.runner = TaskRunner(fn, self.q, self.stop_flag)
        self.runner.start(*args, **kwargs)

    def _tick(self):
        try:
            while True:
                item = self.q.get_nowait()
                kind = item[0]

                if kind == "status":
                    self.set_status(item[1])
                elif kind == "progress":
                    current, total = item[1], item[2]
                    self.progress["maximum"] = total
                    self.progress["value"] = current
                elif kind == "done_ok":
                    self.btn_run.configure(state="normal")
                    self.btn_stop.configure(state="disabled")
                    result = item[1]
                    self.on_done(result)
                elif kind == "done_error":
                    self.btn_run.configure(state="normal")
                    self.btn_stop.configure(state="disabled")
                    self.set_status(f"Error: {item[1]}", "red")
                    messagebox.showerror("Error", item[1])

        except queue.Empty:
            pass

        self.after(120, self._tick)

    def progress_cb(self, current, total):
        """Callback de progreso para pasar a los backends."""
        self.q.put(("progress", current, total))

    def on_done(self, result):
        """Se llama al finalizar con éxito. Puede sobreescribirse."""
        self.set_status("¡Completado!", "green")

    # --- Abstractos ---
    def build_options(self, parent):
        """Construir controles específicos. Implementar en subclases."""
        pass

    def execute(self):
        """Lógica de ejecución. Implementar en subclases."""
        pass

    def get_single_file(self) -> str | None:
        """Obtiene el archivo único (modo single file)."""
        if self.multi_file:
            return None
        path = self.file_entry.get().strip()
        if not path or not os.path.isfile(path):
            messagebox.showwarning("Sin archivo",
                                   "Selecciona un archivo primero.")
            return None
        return path


# ═══════════════════════════════════════════════════════════════════════════════
# PESTAÑAS
# ═══════════════════════════════════════════════════════════════════════════════

class MergeTab(BaseTab):
    def __init__(self, parent, app):
        super().__init__(parent, app, accepted_extensions=["pdf"],
                         multi_file=True)
        self.output_is_folder = False
        self.out_var.set("")
        self.btn_out_browse.configure(
            command=lambda: self._choose_merge_output())

    def _choose_merge_output(self):
        f = filedialog.asksaveasfilename(
            title="Guardar PDF unido como",
            defaultextension=".pdf",
            initialfile="unido.pdf",
            filetypes=[("PDF", "*.pdf")]
        )
        if f:
            self.out_var.set(os.path.abspath(f))

    def build_options(self, parent):
        ttk.Label(parent, text="Los archivos se unirán en el orden de la lista. "
                  "Usa ▲/▼ para reordenar.").pack(anchor="w")

    def execute(self):
        if not self.files:
            messagebox.showwarning("Sin archivos", "Añade PDFs para unir.")
            return

        output = self.out_var.get().strip()
        if not output:
            output = filedialog.asksaveasfilename(
                title="Guardar PDF unido como",
                defaultextension=".pdf",
                initialfile="unido.pdf",
                filetypes=[("PDF", "*.pdf")]
            )
            if not output:
                return
            self.out_var.set(output)

        files = list(self.files)

        def task():
            merge_pdfs(files, output, progress_cb=self.progress_cb)
            size = format_size(os.path.getsize(output))
            return f"PDF unido creado: {os.path.basename(output)} ({size})"

        self.run_task(task)

    def on_done(self, result):
        self.set_status(f"¡Completado! {result}", "green")


class SplitTab(BaseTab):
    def __init__(self, parent, app):
        super().__init__(parent, app, accepted_extensions=["pdf"],
                         multi_file=False)
        self.output_is_folder = True

    def build_options(self, parent):
        self.split_mode = tk.StringVar(value="page_by_page")

        ttk.Radiobutton(parent, text="Página a página",
                        variable=self.split_mode,
                        value="page_by_page").pack(anchor="w")

        n_frame = ttk.Frame(parent)
        n_frame.pack(fill="x", pady=2)
        ttk.Radiobutton(n_frame, text="Cada N páginas:",
                        variable=self.split_mode,
                        value="each_n").pack(side="left")
        self.n_pages_var = tk.StringVar(value="5")
        ttk.Entry(n_frame, textvariable=self.n_pages_var,
                  width=5).pack(side="left", padx=4)

        range_frame = ttk.Frame(parent)
        range_frame.pack(fill="x", pady=2)
        ttk.Radiobutton(range_frame, text="Por rangos:",
                        variable=self.split_mode,
                        value="ranges").pack(side="left")
        self.ranges_var = tk.StringVar(value="1-3, 4-6")
        self.ranges_entry = ttk.Entry(range_frame, textvariable=self.ranges_var)
        self.ranges_entry.pack(side="left", fill="x", expand=True, padx=4)

        # Validación visual de rangos
        self.ranges_entry.bind("<KeyRelease>", self._validate_ranges)
        self.range_feedback = ttk.Label(parent, text="", foreground="gray")
        self.range_feedback.pack(anchor="w")

        prefix_frame = ttk.Frame(parent)
        prefix_frame.pack(fill="x", pady=(4, 0))
        ttk.Label(prefix_frame, text="Prefijo:").pack(side="left")
        self.prefix_var = tk.StringVar(value="parte")
        ttk.Entry(prefix_frame, textvariable=self.prefix_var,
                  width=15).pack(side="left", padx=4)

    def _validate_ranges(self, event=None):
        if self.split_mode.get() != "ranges":
            self.range_feedback.configure(text="", foreground="gray")
            return
        if not self.files:
            return
        try:
            total = get_pdf_page_count(self.files[0])
            if total == 0:
                return
            text = self.ranges_var.get()
            # Validar cada parte
            for part in text.split(","):
                part = part.strip()
                if not part:
                    continue
                if "-" in part:
                    sides = part.split("-", 1)
                    s, e = int(sides[0].strip()), int(sides[1].strip())
                    if s < 1 or e > total or s > e:
                        raise ValueError(f"{part}")
                else:
                    n = int(part.strip())
                    if n < 1 or n > total:
                        raise ValueError(f"{part}")
            self.range_feedback.configure(text="✓ Rangos válidos", foreground="green")
        except (ValueError, IndexError):
            self.range_feedback.configure(text="✗ Rango inválido", foreground="red")

    def execute(self):
        filepath = self.get_single_file()
        if not filepath:
            return

        outdir = self.out_var.get().strip()
        if not outdir:
            outdir = filedialog.askdirectory(title="Carpeta de salida")
            if not outdir:
                return
            self.out_var.set(outdir)

        mode = self.split_mode.get()
        prefix = self.prefix_var.get().strip() or "parte"
        n_pages = int(self.n_pages_var.get() or 5)
        ranges_text = self.ranges_var.get()

        def task():
            created = split_pdf(filepath, mode, outdir, prefix=prefix,
                                n_pages=n_pages, ranges_text=ranges_text,
                                progress_cb=self.progress_cb)
            return f"{len(created)} archivos creados en {outdir}"

        self.run_task(task)

    def on_done(self, result):
        self.set_status(f"¡Completado! {result}", "green")


class ExtractTab(BaseTab):
    def __init__(self, parent, app):
        super().__init__(parent, app, accepted_extensions=["pdf"],
                         multi_file=False)
        self.output_is_folder = False

    def build_options(self, parent):
        ttk.Label(parent, text="Páginas a extraer (ej: 1-3, 5, 8-12):").pack(anchor="w")

        self.ranges_var = tk.StringVar(value="")
        self.ranges_entry = ttk.Entry(parent, textvariable=self.ranges_var,
                                      font=("Consolas", 10))
        self.ranges_entry.pack(fill="x", pady=4)

        self.ranges_entry.bind("<KeyRelease>", self._validate_ranges)
        self.range_feedback = ttk.Label(parent, text="", foreground="gray")
        self.range_feedback.pack(anchor="w")

    def _validate_ranges(self, event=None):
        if not self.files:
            return
        try:
            total = get_pdf_page_count(self.files[0])
            if total == 0:
                return
            text = self.ranges_var.get()
            if not text.strip():
                self.range_feedback.configure(text="", foreground="gray")
                return
            indices = parse_page_ranges(text, total)
            self.range_feedback.configure(
                text=f"✓ {len(indices)} páginas seleccionadas de {total}",
                foreground="green"
            )
        except (ValueError, IndexError) as e:
            self.range_feedback.configure(
                text=f"✗ {e}", foreground="red"
            )

    def execute(self):
        filepath = self.get_single_file()
        if not filepath:
            return

        ranges_text = self.ranges_var.get().strip()
        if not ranges_text:
            messagebox.showwarning("Sin rango", "Especifica las páginas a extraer.")
            return

        output = self.out_var.get().strip()
        if not output:
            output = filedialog.asksaveasfilename(
                title="Guardar extracción como",
                defaultextension=".pdf",
                initialfile="extraido.pdf",
                filetypes=[("PDF", "*.pdf")]
            )
            if not output:
                return
            self.out_var.set(output)

        def task():
            extract_pages(filepath, ranges_text, output,
                          progress_cb=self.progress_cb)
            size = format_size(os.path.getsize(output))
            return f"Extraído: {os.path.basename(output)} ({size})"

        self.run_task(task)

    def on_done(self, result):
        self.set_status(f"¡Completado! {result}", "green")


class RotateTab(BaseTab):
    def __init__(self, parent, app):
        super().__init__(parent, app, accepted_extensions=["pdf"],
                         multi_file=False)
        self.output_is_folder = False

    def build_options(self, parent):
        # Selección de páginas
        sel_frame = ttk.Frame(parent)
        sel_frame.pack(fill="x")

        ttk.Label(sel_frame, text="Páginas:").pack(side="left")
        self.page_sel = tk.StringVar(value="all")
        for val, txt in [("all", "Todas"), ("even", "Pares"),
                         ("odd", "Impares"), ("custom", "Personalizado")]:
            ttk.Radiobutton(sel_frame, text=txt, variable=self.page_sel,
                            value=val).pack(side="left", padx=4)

        range_frame = ttk.Frame(parent)
        range_frame.pack(fill="x", pady=4)
        ttk.Label(range_frame, text="Rango personalizado:").pack(side="left")
        self.custom_range_var = tk.StringVar(value="")
        ttk.Entry(range_frame, textvariable=self.custom_range_var).pack(
            side="left", fill="x", expand=True, padx=4)

        # Grados
        deg_frame = ttk.Frame(parent)
        deg_frame.pack(fill="x", pady=4)
        ttk.Label(deg_frame, text="Rotación:").pack(side="left")
        self.degrees_var = tk.StringVar(value="90")
        for val, txt in [("90", "90° →"), ("180", "180° ↓"), ("270", "270° ←")]:
            ttk.Radiobutton(deg_frame, text=txt, variable=self.degrees_var,
                            value=val).pack(side="left", padx=4)

    def execute(self):
        filepath = self.get_single_file()
        if not filepath:
            return

        output = self.out_var.get().strip()
        if not output:
            output = filedialog.asksaveasfilename(
                title="Guardar PDF rotado como",
                defaultextension=".pdf",
                initialfile="rotado.pdf",
                filetypes=[("PDF", "*.pdf")]
            )
            if not output:
                return
            self.out_var.set(output)

        page_sel = self.page_sel.get()
        degrees = int(self.degrees_var.get())
        custom_range = self.custom_range_var.get()

        def task():
            rotate_pages(filepath, page_sel, degrees, output,
                         ranges_text=custom_range,
                         progress_cb=self.progress_cb)
            return f"Rotado: {os.path.basename(output)}"

        self.run_task(task)

    def on_done(self, result):
        self.set_status(f"¡Completado! {result}", "green")


class ImagesTab(BaseTab):
    def __init__(self, parent, app):
        super().__init__(parent, app,
                         accepted_extensions=["jpg", "jpeg", "png", "bmp",
                                              "tiff", "tif", "webp"],
                         multi_file=True)
        self.output_is_folder = False

    def build_options(self, parent):
        row1 = ttk.Frame(parent)
        row1.pack(fill="x")

        ttk.Label(row1, text="Tamaño de página:").pack(side="left")
        self.page_size_var = tk.StringVar(value="a4")
        for val, txt in [("a4", "A4"), ("letter", "Letter"),
                         ("fit", "Ajustar a imagen")]:
            ttk.Radiobutton(row1, text=txt, variable=self.page_size_var,
                            value=val).pack(side="left", padx=4)

        row2 = ttk.Frame(parent)
        row2.pack(fill="x", pady=4)

        ttk.Label(row2, text="Orientación:").pack(side="left")
        self.orientation_var = tk.StringVar(value="auto")
        for val, txt in [("auto", "Auto"), ("portrait", "Retrato"),
                         ("landscape", "Paisaje")]:
            ttk.Radiobutton(row2, text=txt, variable=self.orientation_var,
                            value=val).pack(side="left", padx=4)

        row3 = ttk.Frame(parent)
        row3.pack(fill="x", pady=2)

        ttk.Label(row3, text="Margen (mm):").pack(side="left")
        self.margin_var = tk.StringVar(value="10")
        ttk.Entry(row3, textvariable=self.margin_var,
                  width=6).pack(side="left", padx=4)

    def execute(self):
        if not self.files:
            messagebox.showwarning("Sin imágenes", "Añade imágenes primero.")
            return

        output = self.out_var.get().strip()
        if not output:
            output = filedialog.asksaveasfilename(
                title="Guardar PDF como",
                defaultextension=".pdf",
                initialfile="imagenes.pdf",
                filetypes=[("PDF", "*.pdf")]
            )
            if not output:
                return
            self.out_var.set(output)

        files = list(self.files)
        page_size = self.page_size_var.get()
        orientation = self.orientation_var.get()
        margin = float(self.margin_var.get() or 10)

        def task():
            images_to_pdf(files, output, page_size=page_size,
                          orientation=orientation, margin_mm=margin,
                          progress_cb=self.progress_cb)
            size = format_size(os.path.getsize(output))
            return f"PDF creado: {os.path.basename(output)} ({size})"

        self.run_task(task)

    def on_done(self, result):
        self.set_status(f"¡Completado! {result}", "green")


class DocxTab(BaseTab):
    def __init__(self, parent, app):
        super().__init__(parent, app,
                         accepted_extensions=["docx", "doc"],
                         multi_file=True)
        self.output_is_folder = True

    def build_options(self, parent):
        self.backend_var = tk.StringVar(value="word" if DOCX2PDF_AVAILABLE else "libreoffice")

        r1 = ttk.Radiobutton(parent,
                              text="Word (docx2pdf) — requiere Microsoft Word",
                              variable=self.backend_var, value="word")
        r1.pack(anchor="w")
        if not DOCX2PDF_AVAILABLE:
            r1.configure(state="disabled")

        r2 = ttk.Radiobutton(parent,
                              text="LibreOffice (headless) — requiere LibreOffice",
                              variable=self.backend_var, value="libreoffice")
        r2.pack(anchor="w")
        if not LIBREOFFICE_PATH:
            r2.configure(state="disabled")

        if not DOCX2PDF_AVAILABLE and not LIBREOFFICE_PATH:
            ttk.Label(parent,
                      text="⚠ No se detectó Word ni LibreOffice. Instala uno para usar esta función.",
                      foreground="red").pack(anchor="w", pady=(4, 0))

    def execute(self):
        if not self.files:
            messagebox.showwarning("Sin archivos", "Añade documentos Word.")
            return

        outdir = self.out_var.get().strip() or None
        backend = self.backend_var.get()
        files = list(self.files)

        def task():
            ok, fail, errors = 0, 0, []
            total = len(files)

            for i, doc in enumerate(files):
                if self.stop_flag.is_set():
                    break

                self.q.put(("status", f"[{i+1}/{total}] Convirtiendo: {os.path.basename(doc)}"))
                try:
                    if backend == "word":
                        convert_with_word(doc, outdir)
                    else:
                        convert_with_libreoffice(doc, outdir)
                    ok += 1
                except Exception as e:
                    fail += 1
                    errors.append(f"- {os.path.basename(doc)}: {e}")

                self.q.put(("progress", i + 1, total))

            if self.stop_flag.is_set():
                return f"Parado. OK: {ok}, Fallos: {fail}"

            if errors:
                error_msg = "\n".join(errors[:10])
                if len(errors) > 10:
                    error_msg += f"\n…y {len(errors)-10} más."
                self.q.put(("status", f"OK: {ok}, Fallos: {fail}"))
                raise RuntimeError(f"Errores:\n{error_msg}")

            return f"Conversión completa. {ok} archivos convertidos."

        self.run_task(task)

    def on_done(self, result):
        self.set_status(f"¡Completado! {result}", "green")


class CompressTab(BaseTab):
    def __init__(self, parent, app):
        super().__init__(parent, app, accepted_extensions=["pdf"],
                         multi_file=True)
        self.output_is_folder = True

    def build_options(self, parent):
        self.quality_var = tk.StringVar(value="ebook")

        ttk.Label(parent, text="Nivel de compresión:").pack(anchor="w")
        for val, txt, desc in [
            ("prepress", "Baja", "mejor calidad, 300 dpi"),
            ("ebook", "Media", "buen balance, 150 dpi"),
            ("screen", "Alta", "menor tamaño, 72 dpi"),
        ]:
            ttk.Radiobutton(parent, text=f"{txt} ({desc})",
                            variable=self.quality_var,
                            value=val).pack(anchor="w")

        self.use_gs = tk.BooleanVar(value=bool(GHOSTSCRIPT_PATH))

        if GHOSTSCRIPT_PATH:
            ttk.Label(parent,
                      text=f"✓ Ghostscript detectado: {GHOSTSCRIPT_PATH}",
                      foreground="green").pack(anchor="w", pady=(6, 0))
            cb = ttk.Checkbutton(parent, text="Usar Ghostscript (mejor compresión)",
                                 variable=self.use_gs)
            cb.pack(anchor="w")
        else:
            ttk.Label(parent,
                      text="⚠ Ghostscript no detectado. Se usará compresión básica (pypdf).",
                      foreground="orange").pack(anchor="w", pady=(6, 0))
            self.use_gs.set(False)

    def execute(self):
        if not self.files:
            messagebox.showwarning("Sin archivos", "Añade PDFs para comprimir.")
            return

        outdir = self.out_var.get().strip()
        if not outdir:
            outdir = filedialog.askdirectory(title="Carpeta de salida")
            if not outdir:
                return
            self.out_var.set(outdir)

        files = list(self.files)
        quality = self.quality_var.get()
        use_gs = self.use_gs.get()
        os.makedirs(outdir, exist_ok=True)

        def task():
            ok, fail = 0, 0
            results = []
            total = len(files)

            for i, f in enumerate(files):
                if self.stop_flag.is_set():
                    break

                name = os.path.basename(f)
                self.q.put(("status", f"[{i+1}/{total}] Comprimiendo: {name}"))

                out_path = os.path.join(outdir, f"comprimido_{name}")
                try:
                    orig_size = os.path.getsize(f)
                    if use_gs:
                        compress_pdf(f, out_path, quality)
                    else:
                        compress_pdf_pypdf(f, out_path)

                    new_size = os.path.getsize(out_path)

                    # Si creció, avisar y usar original
                    if new_size >= orig_size:
                        shutil.copy2(f, out_path)
                        results.append(f"{name}: sin mejora (se mantuvo original)")
                    else:
                        pct = (1 - new_size / orig_size) * 100
                        results.append(
                            f"{name}: {format_size(orig_size)} → {format_size(new_size)} (-{pct:.1f}%)"
                        )
                    ok += 1
                except Exception as e:
                    fail += 1
                    results.append(f"{name}: ERROR — {e}")

                self.q.put(("progress", i + 1, total))

            summary = "\n".join(results)
            return f"OK: {ok}, Fallos: {fail}\n{summary}"

        self.run_task(task)

    def on_done(self, result):
        lines = result.split("\n", 1)
        self.set_status(f"¡Completado! {lines[0]}", "green")


class ProtectTab(BaseTab):
    def __init__(self, parent, app):
        super().__init__(parent, app, accepted_extensions=["pdf"],
                         multi_file=False)
        self.output_is_folder = False

    def build_options(self, parent):
        self.mode_var = tk.StringVar(value="protect")

        mode_frame = ttk.Frame(parent)
        mode_frame.pack(fill="x")
        ttk.Radiobutton(mode_frame, text="Proteger",
                        variable=self.mode_var, value="protect",
                        command=self._toggle_mode).pack(side="left")
        ttk.Radiobutton(mode_frame, text="Desproteger",
                        variable=self.mode_var, value="unprotect",
                        command=self._toggle_mode).pack(side="left", padx=8)

        # Proteger opciones
        self.protect_frame = ttk.Frame(parent)
        self.protect_frame.pack(fill="x", pady=4)

        r1 = ttk.Frame(self.protect_frame)
        r1.pack(fill="x", pady=2)
        ttk.Label(r1, text="Contraseña de usuario:").pack(side="left")
        self.user_pwd_var = tk.StringVar()
        ttk.Entry(r1, textvariable=self.user_pwd_var, show="*",
                  width=25).pack(side="left", padx=4)

        r2 = ttk.Frame(self.protect_frame)
        r2.pack(fill="x", pady=2)
        ttk.Label(r2, text="Contraseña de propietario (opc):").pack(side="left")
        self.owner_pwd_var = tk.StringVar()
        ttk.Entry(r2, textvariable=self.owner_pwd_var, show="*",
                  width=25).pack(side="left", padx=4)

        r3 = ttk.Frame(self.protect_frame)
        r3.pack(fill="x", pady=2)
        ttk.Label(r3, text="Permisos:").pack(side="left")
        self.perm_print = tk.BooleanVar(value=True)
        self.perm_copy = tk.BooleanVar(value=True)
        self.perm_modify = tk.BooleanVar(value=False)
        ttk.Checkbutton(r3, text="Imprimir",
                        variable=self.perm_print).pack(side="left", padx=4)
        ttk.Checkbutton(r3, text="Copiar texto",
                        variable=self.perm_copy).pack(side="left", padx=4)
        ttk.Checkbutton(r3, text="Modificar",
                        variable=self.perm_modify).pack(side="left", padx=4)

        # Desproteger opciones
        self.unprotect_frame = ttk.Frame(parent)

        ur1 = ttk.Frame(self.unprotect_frame)
        ur1.pack(fill="x", pady=2)
        ttk.Label(ur1, text="Contraseña:").pack(side="left")
        self.unlock_pwd_var = tk.StringVar()
        ttk.Entry(ur1, textvariable=self.unlock_pwd_var, show="*",
                  width=25).pack(side="left", padx=4)

    def _toggle_mode(self):
        if self.mode_var.get() == "protect":
            self.unprotect_frame.pack_forget()
            self.protect_frame.pack(fill="x", pady=4)
        else:
            self.protect_frame.pack_forget()
            self.unprotect_frame.pack(fill="x", pady=4)

    def execute(self):
        filepath = self.get_single_file()
        if not filepath:
            return

        output = self.out_var.get().strip()
        if not output:
            mode_str = "protegido" if self.mode_var.get() == "protect" else "desprotegido"
            output = filedialog.asksaveasfilename(
                title="Guardar como",
                defaultextension=".pdf",
                initialfile=f"{mode_str}.pdf",
                filetypes=[("PDF", "*.pdf")]
            )
            if not output:
                return
            self.out_var.set(output)

        mode = self.mode_var.get()

        if mode == "protect":
            user_pwd = self.user_pwd_var.get()
            if not user_pwd:
                messagebox.showwarning("Sin contraseña",
                                       "Introduce una contraseña de usuario.")
                return

            owner_pwd = self.owner_pwd_var.get()
            permissions = {
                "print": self.perm_print.get(),
                "copy": self.perm_copy.get(),
                "modify": self.perm_modify.get(),
            }

            def task():
                protect_pdf(filepath, output, user_pwd,
                            owner_password=owner_pwd,
                            permissions=permissions,
                            progress_cb=self.progress_cb)
                return f"PDF protegido: {os.path.basename(output)}"

        else:
            pwd = self.unlock_pwd_var.get()
            if not pwd:
                messagebox.showwarning("Sin contraseña",
                                       "Introduce la contraseña del PDF.")
                return

            def task():
                unprotect_pdf(filepath, output, pwd,
                              progress_cb=self.progress_cb)
                return f"PDF desprotegido: {os.path.basename(output)}"

        self.run_task(task)

    def on_done(self, result):
        self.set_status(f"¡Completado! {result}", "green")


class WatermarkTab(BaseTab):
    """Pestaña para marca de agua y metadatos."""

    def __init__(self, parent, app):
        super().__init__(parent, app, accepted_extensions=["pdf"],
                         multi_file=False)
        self.output_is_folder = False

    def build_options(self, parent):
        self.action_var = tk.StringVar(value="watermark")

        mode_frame = ttk.Frame(parent)
        mode_frame.pack(fill="x")
        ttk.Radiobutton(mode_frame, text="Marca de agua",
                        variable=self.action_var, value="watermark",
                        command=self._toggle_mode).pack(side="left")
        ttk.Radiobutton(mode_frame, text="Metadatos",
                        variable=self.action_var, value="metadata",
                        command=self._toggle_mode).pack(side="left", padx=8)

        # Watermark frame
        self.wm_frame = ttk.Frame(parent)
        self.wm_frame.pack(fill="x", pady=4)

        r1 = ttk.Frame(self.wm_frame)
        r1.pack(fill="x", pady=2)
        ttk.Label(r1, text="Texto:").pack(side="left")
        self.wm_text_var = tk.StringVar(value="BORRADOR")
        ttk.Entry(r1, textvariable=self.wm_text_var,
                  width=30).pack(side="left", padx=4)

        r2 = ttk.Frame(self.wm_frame)
        r2.pack(fill="x", pady=2)
        ttk.Label(r2, text="Tamaño fuente:").pack(side="left")
        self.wm_fontsize_var = tk.StringVar(value="50")
        ttk.Entry(r2, textvariable=self.wm_fontsize_var,
                  width=5).pack(side="left", padx=4)

        ttk.Label(r2, text="Opacidad (0-1):").pack(side="left", padx=(8, 0))
        self.wm_opacity_var = tk.StringVar(value="0.3")
        ttk.Entry(r2, textvariable=self.wm_opacity_var,
                  width=5).pack(side="left", padx=4)

        ttk.Label(r2, text="Rotación:").pack(side="left", padx=(8, 0))
        self.wm_rotation_var = tk.StringVar(value="45")
        ttk.Entry(r2, textvariable=self.wm_rotation_var,
                  width=5).pack(side="left", padx=4)

        # Metadata frame
        self.meta_frame = ttk.Frame(parent)

        for attr, label in [("meta_title", "Título:"),
                            ("meta_author", "Autor:"),
                            ("meta_subject", "Asunto:"),
                            ("meta_keywords", "Palabras clave:")]:
            row = ttk.Frame(self.meta_frame)
            row.pack(fill="x", pady=2)
            ttk.Label(row, text=label, width=16).pack(side="left")
            var = tk.StringVar()
            setattr(self, f"{attr}_var", var)
            ttk.Entry(row, textvariable=var).pack(side="left", fill="x",
                                                   expand=True, padx=4)

    def _toggle_mode(self):
        if self.action_var.get() == "watermark":
            self.meta_frame.pack_forget()
            self.wm_frame.pack(fill="x", pady=4)
        else:
            self.wm_frame.pack_forget()
            self.meta_frame.pack(fill="x", pady=4)

    def execute(self):
        filepath = self.get_single_file()
        if not filepath:
            return

        output = self.out_var.get().strip()
        if not output:
            output = filedialog.asksaveasfilename(
                title="Guardar como",
                defaultextension=".pdf",
                filetypes=[("PDF", "*.pdf")]
            )
            if not output:
                return
            self.out_var.set(output)

        action = self.action_var.get()

        if action == "watermark":
            text = self.wm_text_var.get()
            font_size = int(self.wm_fontsize_var.get() or 50)
            opacity = float(self.wm_opacity_var.get() or 0.3)
            rotation = int(self.wm_rotation_var.get() or 45)

            def task():
                add_watermark_text(filepath, output, text,
                                   font_size=font_size, opacity=opacity,
                                   rotation=rotation,
                                   progress_cb=self.progress_cb)
                return f"Marca de agua añadida: {os.path.basename(output)}"

        else:
            metadata = {}
            if self.meta_title_var.get():
                metadata["/Title"] = self.meta_title_var.get()
            if self.meta_author_var.get():
                metadata["/Author"] = self.meta_author_var.get()
            if self.meta_subject_var.get():
                metadata["/Subject"] = self.meta_subject_var.get()
            if self.meta_keywords_var.get():
                metadata["/Keywords"] = self.meta_keywords_var.get()

            if not metadata:
                messagebox.showwarning("Sin datos",
                                       "Rellena al menos un campo de metadatos.")
                return

            def task():
                set_pdf_metadata(filepath, output, metadata,
                                 progress_cb=self.progress_cb)
                return f"Metadatos actualizados: {os.path.basename(output)}"

        self.run_task(task)

    def on_done(self, result):
        self.set_status(f"¡Completado! {result}", "green")


# ═══════════════════════════════════════════════════════════════════════════════
# APLICACIÓN PRINCIPAL
# ═══════════════════════════════════════════════════════════════════════════════

class App:
    DARK_BG = "#2b2b2b"
    DARK_FG = "#e0e0e0"
    DARK_SELECT = "#404040"
    DARK_ACCENT = "#3c7dc4"

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("PDF Toolkit")
        self.root.geometry("920x680")
        self.root.minsize(750, 550)

        self.config = load_config()
        self.dark_mode = self.config.get("dark_mode", False)

        self._setup_style()
        self._build_ui()
        self._apply_theme()

        # Restaurar última posición
        geom = self.config.get("geometry")
        if geom:
            try:
                self.root.geometry(geom)
            except Exception:
                pass

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    def _setup_style(self):
        self.style = ttk.Style()
        try:
            self.style.theme_use("clam")
        except Exception:
            pass

    def _build_ui(self):
        # Menu bar
        menubar = tk.Menu(self.root)
        self.root.configure(menu=menubar)

        view_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Ver", menu=view_menu)
        self.dark_mode_var = tk.BooleanVar(value=self.dark_mode)
        view_menu.add_checkbutton(label="Modo oscuro",
                                  variable=self.dark_mode_var,
                                  command=self._toggle_dark_mode)

        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Ayuda", menu=help_menu)
        help_menu.add_command(label="Acerca de…", command=self._show_about)
        help_menu.add_command(label="Dependencias…", command=self._show_deps)

        # Main container
        self.main_frame = ttk.Frame(self.root, padding=5)
        self.main_frame.pack(fill="both", expand=True)

        # Welcome banner
        self.welcome = ttk.Label(
            self.main_frame,
            text="PDF Toolkit — Selecciona una pestaña para comenzar",
            font=("Segoe UI", 11, "bold") if sys.platform == "win32" else ("Helvetica", 12, "bold"),
            foreground="#3c7dc4"
        )
        self.welcome.pack(pady=(0, 6))

        # Notebook
        self.notebook = ttk.Notebook(self.main_frame)
        self.notebook.pack(fill="both", expand=True)

        # Crear pestañas
        self.tabs = {}
        tab_defs = [
            ("Unir", MergeTab, PYPDF_AVAILABLE),
            ("Dividir", SplitTab, PYPDF_AVAILABLE),
            ("Extraer", ExtractTab, PYPDF_AVAILABLE),
            ("Rotar", RotateTab, PYPDF_AVAILABLE),
            ("Imágenes→PDF", ImagesTab, PILLOW_AVAILABLE),
            ("DOCX→PDF", DocxTab, DOCX2PDF_AVAILABLE or bool(LIBREOFFICE_PATH)),
            ("Comprimir", CompressTab, PYPDF_AVAILABLE),
            ("Proteger", ProtectTab, PYPDF_AVAILABLE),
            ("Marca/Meta", WatermarkTab, PYPDF_AVAILABLE),
        ]

        for name, cls, available in tab_defs:
            tab = cls(self.notebook, self)
            self.notebook.add(tab, text=name)
            self.tabs[name] = tab

            if not available:
                # Deshabilitar pestaña visualmente
                idx = self.notebook.index("end") - 1
                self.notebook.tab(idx, state="disabled")

        # Restaurar última pestaña
        last_tab = self.config.get("last_tab", 0)
        try:
            self.notebook.select(last_tab)
        except Exception:
            pass

    def _toggle_dark_mode(self):
        self.dark_mode = self.dark_mode_var.get()
        self._apply_theme()

    def _apply_theme(self):
        if self.dark_mode:
            self.style.configure(".", background=self.DARK_BG,
                                 foreground=self.DARK_FG,
                                 fieldbackground=self.DARK_BG)
            self.style.configure("TFrame", background=self.DARK_BG)
            self.style.configure("TLabel", background=self.DARK_BG,
                                 foreground=self.DARK_FG)
            self.style.configure("TLabelframe", background=self.DARK_BG,
                                 foreground=self.DARK_FG)
            self.style.configure("TLabelframe.Label", background=self.DARK_BG,
                                 foreground=self.DARK_FG)
            self.style.configure("TNotebook", background=self.DARK_BG)
            self.style.configure("TNotebook.Tab", background="#3a3a3a",
                                 foreground=self.DARK_FG)
            self.style.map("TNotebook.Tab",
                           background=[("selected", self.DARK_ACCENT)])
            self.style.configure("TButton", background="#3a3a3a",
                                 foreground=self.DARK_FG)
            self.style.configure("TRadiobutton", background=self.DARK_BG,
                                 foreground=self.DARK_FG)
            self.style.configure("TCheckbutton", background=self.DARK_BG,
                                 foreground=self.DARK_FG)
            self.style.configure("TEntry", fieldbackground="#3a3a3a",
                                 foreground=self.DARK_FG)
            self.root.configure(bg=self.DARK_BG)

            # Listboxes
            for tab in self.tabs.values():
                if hasattr(tab, "listbox") and tab.listbox:
                    tab.listbox.configure(
                        bg="#3a3a3a", fg=self.DARK_FG,
                        selectbackground=self.DARK_ACCENT,
                        selectforeground="white"
                    )
        else:
            self.style.theme_use("clam")
            self.root.configure(bg="")

            for tab in self.tabs.values():
                if hasattr(tab, "listbox") and tab.listbox:
                    tab.listbox.configure(
                        bg="white", fg="black",
                        selectbackground="#0078d7",
                        selectforeground="white"
                    )

    def _show_about(self):
        messagebox.showinfo(
            "Acerca de PDF Toolkit",
            "PDF Toolkit v1.0\n\n"
            "Herramienta completa de manipulación de PDFs.\n\n"
            "Funciones: Unir, Dividir, Extraer, Rotar,\n"
            "Imágenes→PDF, DOCX→PDF, Comprimir,\n"
            "Proteger/Desproteger, Marca de agua, Metadatos.\n\n"
            "Basado en pypdf, Pillow y Tkinter."
        )

    def _show_deps(self):
        deps = [
            f"pypdf: {'✓' if PYPDF_AVAILABLE else '✗ (pip install pypdf)'}",
            f"Pillow: {'✓' if PILLOW_AVAILABLE else '✗ (pip install Pillow)'}",
            f"tkinterdnd2: {'✓' if DND_AVAILABLE else '✗ (pip install tkinterdnd2)'}",
            f"docx2pdf: {'✓' if DOCX2PDF_AVAILABLE else '✗ (pip install docx2pdf)'}",
            f"LibreOffice: {'✓ ' + (LIBREOFFICE_PATH or '') if LIBREOFFICE_PATH else '✗ (no en PATH)'}",
            f"Ghostscript: {'✓ ' + (GHOSTSCRIPT_PATH or '') if GHOSTSCRIPT_PATH else '✗ (no en PATH)'}",
        ]
        messagebox.showinfo("Dependencias", "\n".join(deps))

    def _on_close(self):
        self.config["dark_mode"] = self.dark_mode
        self.config["geometry"] = self.root.geometry()
        try:
            self.config["last_tab"] = self.notebook.index(self.notebook.select())
        except Exception:
            pass
        save_config(self.config)
        self.root.destroy()


# ═══════════════════════════════════════════════════════════════════════════════
# LÍNEA DE COMANDOS
# ═══════════════════════════════════════════════════════════════════════════════

def cli_main():
    parser = argparse.ArgumentParser(
        description="PDF Toolkit — Manipulación de PDFs desde línea de comandos"
    )
    sub = parser.add_subparsers(dest="command")

    # Merge
    p_merge = sub.add_parser("merge", help="Unir PDFs")
    p_merge.add_argument("files", nargs="+", help="PDFs a unir")
    p_merge.add_argument("-o", "--output", required=True, help="Archivo de salida")

    # Split
    p_split = sub.add_parser("split", help="Dividir un PDF")
    p_split.add_argument("file", help="PDF a dividir")
    p_split.add_argument("-m", "--mode", choices=["page_by_page", "each_n", "ranges"],
                         default="page_by_page")
    p_split.add_argument("-n", "--n-pages", type=int, default=1)
    p_split.add_argument("-r", "--ranges", default="")
    p_split.add_argument("-o", "--outdir", required=True)
    p_split.add_argument("-p", "--prefix", default="parte")

    # Extract
    p_extract = sub.add_parser("extract", help="Extraer páginas")
    p_extract.add_argument("file", help="PDF fuente")
    p_extract.add_argument("-r", "--ranges", required=True, help="Rangos: 1-3,5,8")
    p_extract.add_argument("-o", "--output", required=True)

    # Rotate
    p_rotate = sub.add_parser("rotate", help="Rotar páginas")
    p_rotate.add_argument("file", help="PDF fuente")
    p_rotate.add_argument("-d", "--degrees", type=int, choices=[90, 180, 270],
                          required=True)
    p_rotate.add_argument("-s", "--selection", choices=["all", "even", "odd", "custom"],
                          default="all")
    p_rotate.add_argument("-r", "--ranges", default="")
    p_rotate.add_argument("-o", "--output", required=True)

    # Images to PDF
    p_img = sub.add_parser("images", help="Imágenes a PDF")
    p_img.add_argument("files", nargs="+", help="Imágenes")
    p_img.add_argument("-o", "--output", required=True)
    p_img.add_argument("--page-size", choices=["a4", "letter", "fit"], default="a4")
    p_img.add_argument("--orientation", choices=["auto", "portrait", "landscape"],
                       default="auto")
    p_img.add_argument("--margin", type=float, default=10)

    # Compress
    p_comp = sub.add_parser("compress", help="Comprimir PDF")
    p_comp.add_argument("file", help="PDF a comprimir")
    p_comp.add_argument("-o", "--output", required=True)
    p_comp.add_argument("-q", "--quality", choices=["prepress", "ebook", "screen"],
                        default="ebook")

    # Protect
    p_prot = sub.add_parser("protect", help="Proteger PDF")
    p_prot.add_argument("file", help="PDF a proteger")
    p_prot.add_argument("-o", "--output", required=True)
    p_prot.add_argument("--user-password", required=True)
    p_prot.add_argument("--owner-password", default="")

    # Unprotect
    p_unprot = sub.add_parser("unprotect", help="Desproteger PDF")
    p_unprot.add_argument("file", help="PDF protegido")
    p_unprot.add_argument("-o", "--output", required=True)
    p_unprot.add_argument("--password", required=True)

    # Metadata
    p_meta = sub.add_parser("metadata", help="Establecer metadatos")
    p_meta.add_argument("file", help="PDF")
    p_meta.add_argument("-o", "--output", required=True)
    p_meta.add_argument("--title", default="")
    p_meta.add_argument("--author", default="")
    p_meta.add_argument("--subject", default="")

    args = parser.parse_args()

    if not args.command:
        return False  # No CLI command, launch GUI

    # Execute CLI command
    def progress(c, t):
        print(f"  [{c}/{t}]", end="\r")

    try:
        if args.command == "merge":
            merge_pdfs(args.files, args.output, progress_cb=progress)
            print(f"✓ PDF unido: {args.output}")

        elif args.command == "split":
            created = split_pdf(args.file, args.mode, args.outdir,
                                prefix=args.prefix, n_pages=args.n_pages,
                                ranges_text=args.ranges, progress_cb=progress)
            print(f"✓ {len(created)} archivos creados")

        elif args.command == "extract":
            extract_pages(args.file, args.ranges, args.output,
                          progress_cb=progress)
            print(f"✓ Páginas extraídas: {args.output}")

        elif args.command == "rotate":
            rotate_pages(args.file, args.selection, args.degrees, args.output,
                         ranges_text=args.ranges, progress_cb=progress)
            print(f"✓ PDF rotado: {args.output}")

        elif args.command == "images":
            images_to_pdf(args.files, args.output, page_size=args.page_size,
                          orientation=args.orientation, margin_mm=args.margin,
                          progress_cb=progress)
            print(f"✓ PDF creado: {args.output}")

        elif args.command == "compress":
            orig = os.path.getsize(args.file)
            if GHOSTSCRIPT_PATH:
                compress_pdf(args.file, args.output, args.quality,
                             progress_cb=progress)
            else:
                compress_pdf_pypdf(args.file, args.output, progress_cb=progress)
            new = os.path.getsize(args.output)
            pct = (1 - new / orig) * 100 if orig else 0
            print(f"✓ Comprimido: {format_size(orig)} → {format_size(new)} (-{pct:.1f}%)")

        elif args.command == "protect":
            protect_pdf(args.file, args.output, args.user_password,
                        owner_password=args.owner_password,
                        progress_cb=progress)
            print(f"✓ PDF protegido: {args.output}")

        elif args.command == "unprotect":
            unprotect_pdf(args.file, args.output, args.password,
                          progress_cb=progress)
            print(f"✓ PDF desprotegido: {args.output}")

        elif args.command == "metadata":
            meta = {}
            if args.title:
                meta["/Title"] = args.title
            if args.author:
                meta["/Author"] = args.author
            if args.subject:
                meta["/Subject"] = args.subject
            set_pdf_metadata(args.file, args.output, meta, progress_cb=progress)
            print(f"✓ Metadatos actualizados: {args.output}")

    except Exception as e:
        print(f"✗ Error: {e}", file=sys.stderr)
        sys.exit(1)

    return True


# ═══════════════════════════════════════════════════════════════════════════════
# PUNTO DE ENTRADA
# ═══════════════════════════════════════════════════════════════════════════════

def main():
    # Si hay argumentos CLI, intentar ejecutar en modo CLI
    if len(sys.argv) > 1:
        if cli_main():
            return

    # Lanzar GUI
    if DND_AVAILABLE:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()

    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
