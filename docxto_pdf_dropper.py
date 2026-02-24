import os
import sys
import threading
import queue
import subprocess
import shutil
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# --- Drag & Drop (tkinterdnd2) ---
DND_AVAILABLE = True
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
except Exception:
    DND_AVAILABLE = False
    TkinterDnD = None
    DND_FILES = None

# --- Conversion backends ---
def convert_with_word(docx_path: str, outdir: str | None) -> None:
    """
    Uses docx2pdf -> Microsoft Word COM (Windows/macOS).
    If outdir is None: output is placed next to input file.
    If outdir is a directory: output is placed there with same basename.
    """
    from docx2pdf import convert

    # docx2pdf uses COM on Windows; make it safer in worker threads
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
            convert(docx_path, outdir)
        else:
            convert(docx_path)
    finally:
        if sys.platform.startswith("win") and pythoncom is not None:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass


def convert_with_libreoffice(docx_path: str, outdir: str | None) -> None:
    """
    Uses LibreOffice headless conversion: soffice --headless --convert-to pdf ...
    If outdir is None: output is placed next to input file.
    """
    target_dir = outdir or os.path.dirname(docx_path)
    os.makedirs(target_dir, exist_ok=True)

    # Try to find soffice in PATH
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        raise RuntimeError(
            "No encuentro 'soffice' (LibreOffice) en el PATH. "
            "Instala LibreOffice o añade soffice.exe al PATH."
        )

    cmd = [
        soffice,
        "--headless",
        "--nologo",
        "--nolockcheck",
        "--norestore",
        "--convert-to",
        "pdf",
        "--outdir",
        target_dir,
        docx_path,
    ]

    # LibreOffice a veces devuelve 0 aunque falle; captura stderr por si acaso.
    p = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    if p.returncode != 0:
        raise RuntimeError(p.stderr.strip() or f"LibreOffice falló (code {p.returncode}).")


# --- Helpers ---
def is_docx(path: str) -> bool:
    ext = os.path.splitext(path)[1].lower()
    return ext in [".docx", ".doc"]  # si quieres SOLO docx, quita ".doc"

def collect_docs(paths: list[str]) -> list[str]:
    out: list[str] = []
    for p in paths:
        p = p.strip().strip('"')
        if not p:
            continue
        if os.path.isdir(p):
            for root, _, files in os.walk(p):
                for f in files:
                    fp = os.path.join(root, f)
                    if is_docx(fp):
                        out.append(os.path.abspath(fp))
        else:
            if os.path.isfile(p) and is_docx(p):
                out.append(os.path.abspath(p))
    return out

def expected_pdf_path(doc_path: str, outdir: str | None) -> str:
    base = os.path.splitext(os.path.basename(doc_path))[0] + ".pdf"
    if outdir:
        return os.path.join(outdir, base)
    return os.path.join(os.path.dirname(doc_path), base)


# --- UI App ---
class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("DOCX → PDF (en masa)")
        self.root.geometry("840x520")

        self.files: list[str] = []
        self.file_set: set[str] = set()

        self.outdir: str | None = None
        self.backend = tk.StringVar(value="word")  # "word" or "libreoffice"

        self.q: queue.Queue = queue.Queue()
        self.worker: threading.Thread | None = None
        self.stop_flag = threading.Event()

        self._build_ui()
        self._tick()

    def _build_ui(self):
        top = ttk.Frame(self.root, padding=10)
        top.pack(fill="both", expand=True)

        # Drop zone / list
        list_frame = ttk.LabelFrame(top, text="Arrastra aquí archivos o carpetas (.doc/.docx)", padding=10)
        list_frame.pack(fill="both", expand=True)

        self.listbox = tk.Listbox(list_frame, selectmode=tk.EXTENDED)
        self.listbox.pack(side="left", fill="both", expand=True)

        sb = ttk.Scrollbar(list_frame, orient="vertical", command=self.listbox.yview)
        sb.pack(side="right", fill="y")
        self.listbox.configure(yscrollcommand=sb.set)

        if DND_AVAILABLE:
            try:
                # Root must be TkinterDnD.Tk in that case (handled outside)
                self.listbox.drop_target_register(DND_FILES)
                self.listbox.dnd_bind("<<Drop>>", self._on_drop)
            except Exception:
                pass

        # Controls
        controls = ttk.Frame(top, padding=(0, 10, 0, 0))
        controls.pack(fill="x")

        ttk.Button(controls, text="Añadir archivos…", command=self._add_files_dialog).pack(side="left")
        ttk.Button(controls, text="Quitar selección", command=self._remove_selected).pack(side="left", padx=6)
        ttk.Button(controls, text="Limpiar", command=self._clear).pack(side="left")

        # Output folder
        out_frame = ttk.Frame(controls)
        out_frame.pack(side="right")

        self.out_label = ttk.Label(out_frame, text="Salida: (misma carpeta)")
        self.out_label.pack(side="left", padx=(0, 6))
        ttk.Button(out_frame, text="Elegir carpeta…", command=self._choose_outdir).pack(side="left")
        ttk.Button(out_frame, text="Reset", command=self._reset_outdir).pack(side="left", padx=(6, 0))

        # Backend selection
        backend_frame = ttk.LabelFrame(top, text="Motor de conversión", padding=10)
        backend_frame.pack(fill="x")

        ttk.Radiobutton(
            backend_frame, text="Word (docx2pdf) — requiere Microsoft Word",
            variable=self.backend, value="word"
        ).pack(anchor="w")
        ttk.Radiobutton(
            backend_frame, text="LibreOffice (headless) — requiere LibreOffice/soffice",
            variable=self.backend, value="libreoffice"
        ).pack(anchor="w")

        # Bottom: progress + buttons
        bottom = ttk.Frame(top, padding=(0, 10, 0, 0))
        bottom.pack(fill="x")

        self.progress = ttk.Progressbar(bottom, mode="determinate")
        self.progress.pack(fill="x")

        self.status = ttk.Label(bottom, text="Listo.")
        self.status.pack(anchor="w", pady=(6, 0))

        btns = ttk.Frame(bottom)
        btns.pack(fill="x", pady=(10, 0))

        self.btn_convert = ttk.Button(btns, text="Convertir a PDF", command=self._start)
        self.btn_convert.pack(side="left")

        self.btn_stop = ttk.Button(btns, text="Parar", command=self._stop, state="disabled")
        self.btn_stop.pack(side="left", padx=6)

        ttk.Label(btns, text="Tip: puedes arrastrar carpetas enteras.").pack(side="right")

        if not DND_AVAILABLE:
            messagebox.showwarning(
                "Drag&Drop no disponible",
                "No puedo importar tkinterdnd2.\n"
                "Instala con:  py -m pip install tkinterdnd2\n"
                "Mientras tanto, usa 'Añadir archivos…'"
            )

    def _on_drop(self, event):
        try:
            # Tk gives a Tcl list; splitlist handles braces/spaces correctly.
            paths = list(self.root.tk.splitlist(event.data))
        except Exception:
            paths = [event.data]

        self._add_paths(paths)

    def _add_files_dialog(self):
        files = filedialog.askopenfilenames(
            title="Selecciona archivos Word",
            filetypes=[("Word", "*.docx *.doc"), ("Todos", "*.*")]
        )
        if files:
            self._add_paths(list(files))

    def _add_paths(self, paths: list[str]):
        docs = collect_docs(paths)
        added = 0
        for p in docs:
            if p not in self.file_set:
                self.file_set.add(p)
                self.files.append(p)
                self.listbox.insert(tk.END, p)
                added += 1

        if added:
            self.status.configure(text=f"Añadidos: {added}. Total: {len(self.files)}")
        else:
            self.status.configure(text="No he encontrado .doc/.docx nuevos en lo que has soltado.")

    def _remove_selected(self):
        sel = list(self.listbox.curselection())
        if not sel:
            return
        # remove from end to start to keep indices valid
        for idx in reversed(sel):
            p = self.listbox.get(idx)
            self.listbox.delete(idx)
            self.file_set.discard(p)
        # rebuild list
        self.files = list(self.listbox.get(0, tk.END))
        self.status.configure(text=f"Total: {len(self.files)}")

    def _clear(self):
        self.listbox.delete(0, tk.END)
        self.files.clear()
        self.file_set.clear()
        self.status.configure(text="Listo (lista vacía).")

    def _choose_outdir(self):
        d = filedialog.askdirectory(title="Carpeta de salida")
        if d:
            self.outdir = os.path.abspath(d)
            self.out_label.configure(text=f"Salida: {self.outdir}")

    def _reset_outdir(self):
        self.outdir = None
        self.out_label.configure(text="Salida: (misma carpeta)")

    def _start(self):
        if self.worker and self.worker.is_alive():
            return
        if not self.files:
            messagebox.showinfo("Nada que convertir", "Arrastra archivos o añade con el botón.")
            return

        self.stop_flag.clear()
        self.progress["value"] = 0
        self.progress["maximum"] = len(self.files)
        self.btn_convert.configure(state="disabled")
        self.btn_stop.configure(state="normal")

        self.worker = threading.Thread(target=self._run_conversion, daemon=True)
        self.worker.start()

    def _stop(self):
        self.stop_flag.set()
        self.status.configure(text="Parando… (deja terminar el archivo actual)")

    def _run_conversion(self):
        total = len(self.files)
        ok = 0
        fail = 0
        errors: list[str] = []

        backend = self.backend.get().strip()

        for i, doc in enumerate(self.files, start=1):
            if self.stop_flag.is_set():
                break

            pdf_path = expected_pdf_path(doc, self.outdir)
            self.q.put(("status", f"[{i}/{total}] Convirtiendo: {os.path.basename(doc)}"))
            try:
                if backend == "word":
                    convert_with_word(doc, self.outdir)
                else:
                    convert_with_libreoffice(doc, self.outdir)

                ok += 1
            except Exception as e:
                fail += 1
                errors.append(f"- {doc}\n  {type(e).__name__}: {e}")

            self.q.put(("progress", i))

        # finish
        self.q.put(("done", ok, fail, errors))

    def _tick(self):
        try:
            while True:
                item = self.q.get_nowait()
                kind = item[0]

                if kind == "status":
                    self.status.configure(text=item[1])

                elif kind == "progress":
                    self.progress["value"] = item[1]

                elif kind == "done":
                    ok, fail, errors = item[1], item[2], item[3]
                    self.btn_convert.configure(state="normal")
                    self.btn_stop.configure(state="disabled")

                    if self.stop_flag.is_set():
                        self.status.configure(text=f"Parado. OK: {ok}, Fallos: {fail}")
                    else:
                        self.status.configure(text=f"Terminado. OK: {ok}, Fallos: {fail}")

                    if fail:
                        msg = "Algunos archivos fallaron:\n\n" + "\n\n".join(errors[:10])
                        if len(errors) > 10:
                            msg += f"\n\n…y {len(errors)-10} más."
                        messagebox.showerror("Errores en la conversión", msg)
        except queue.Empty:
            pass

        self.root.after(120, self._tick)


def main():
    # Use TkinterDnD root if available for drag&drop
    if DND_AVAILABLE:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()

    # Nice-ish default style
    try:
        ttk.Style().theme_use("clam")
    except Exception:
        pass

    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
