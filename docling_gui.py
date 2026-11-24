"""DocuMark Transformer: convert PDF, DOCX, PPTX and HTML files to Markdown with Docling.

Startup design (why it opens fast):
  * Only PySide6 is imported before the window appears. Docling (and torch under it)
    is imported in a background thread once the window has painted.
  * The UI is usable immediately. A file picked while Docling is still loading is
    queued and converts as soon as the engine is ready.
  * One DocumentConverter lives for the whole session, and its PDF models are loaded
    in the background right after start, so the first PDF does not pay for them.

Set DOCUMARK_TIMING=1 to print startup and conversion timings to the console.
"""
import os
import sys
import time

_T0 = time.perf_counter()
TIMING = bool(os.environ.get("DOCUMARK_TIMING"))


def mark(label: str) -> None:
    if TIMING:
        print(f"[timing] {time.perf_counter() - _T0:6.2f}s  {label}", flush=True)


# Hugging Face Hub symlinks need admin rights or Developer Mode on Windows; copy instead.
# Must be set before Docling (and huggingface_hub) is imported.
if sys.platform == "win32":
    os.environ.setdefault("HF_HUB_DISABLE_SYMLINKS", "1")
    os.environ.setdefault("HF_HUB_DISABLE_SYMLINKS_WARNING", "1")

from PySide6.QtCore import QObject, Qt, QThread, QTimer, Signal, Slot
from PySide6.QtGui import QDragEnterEvent, QDropEvent, QFont
from PySide6.QtWidgets import (
    QApplication, QFileDialog, QHBoxLayout, QLabel, QMainWindow, QMessageBox,
    QProgressDialog, QPushButton, QTextEdit, QVBoxLayout, QWidget,
)

mark("Qt imported")

OPEN_FILTERS = (
    "All Supported Files (*.pdf *.docx *.pptx *.html *.htm);;"
    "PDF Files (*.pdf);;Word Documents (*.docx);;PowerPoint Files (*.pptx);;"
    "HTML Files (*.html *.htm);;All Files (*)"
)
SAVE_FILTERS = "Markdown Files (*.md);;Text Files (*.txt);;All Files (*)"


# ==============================================================
# Conversion engine: lives in its own thread for the whole session
# ==============================================================
class Engine(QObject):
    """Owns the Docling converter. All slots run in the engine thread, one at a time,
    so jobs submitted before loading finishes simply wait in the thread's queue."""

    ready = Signal()                       # converter usable (DOCX/PPTX/HTML right away)
    load_failed = Signal(str, str)         # error type, message
    converted = Signal(int, str)           # job id, markdown
    failed = Signal(int, str, str)         # job id, error type, message

    def __init__(self):
        super().__init__()
        self._converter = None

    @Slot()
    def load(self):
        try:
            from docling.datamodel.base_models import InputFormat
            from docling.document_converter import DocumentConverter
            mark("docling imported")
            self._converter = DocumentConverter()
        except Exception as e:  # ImportError, or a broken install
            self.load_failed.emit(type(e).__name__, str(e))
            return
        self.ready.emit()
        # Queue the warm-up behind any file the user already picked while Docling
        # was loading, so that file is converted first.
        QTimer.singleShot(0, lambda: self._warm_up(InputFormat.PDF))

    def _warm_up(self, pdf_format):
        try:
            # Load layout/table/OCR models now rather than on the first PDF.
            # On the very first launch this also downloads them (one time only).
            self._converter.initialize_pipeline(pdf_format)
            mark("PDF models loaded")
        except Exception as e:
            # Not fatal: the next PDF conversion retries and reports a real error.
            print(f"PDF model warm-up failed: {type(e).__name__}: {e}", file=sys.stderr)

    @Slot(int, str)
    def convert(self, job_id: int, path: str):
        if self._converter is None:
            self.failed.emit(job_id, "NotReady", "Docling failed to load.")
            return
        try:
            t = time.perf_counter()
            doc = self._converter.convert(path).document
            markdown = doc.export_to_markdown(image_placeholder="")
            mark(f"converted {os.path.basename(path)} in {time.perf_counter() - t:.2f}s")
            self.converted.emit(job_id, markdown)
        except Exception as e:
            self.failed.emit(job_id, type(e).__name__, str(e))


# ==============================================================
# Main window
# ==============================================================
class MarkdownConverterApp(QMainWindow):
    submit = Signal(int, str)   # -> Engine.convert

    READY_TEXT = "Ready. Select a file or drag it here."

    def __init__(self):
        super().__init__()
        self.setWindowTitle("DocuMark Transformer - Convert Documents to Markdown")
        self.setGeometry(100, 100, 800, 650)
        self.setAcceptDrops(True)

        self.engine_state = "loading"    # loading | ready | failed
        self.engine_error = ""
        self.job_id = 0                  # id of the job the UI is waiting for (0 = none)
        self.cancelled = False
        self.last_processed_file = None
        self.baseline_status = "Loading Docling in the background. You can pick a file now."
        self.progress = None

        self._build_ui()
        self._start_engine()

    # ---------------- UI ----------------
    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setContentsMargins(20, 20, 20, 15)
        layout.setSpacing(15)

        title = QLabel("✨ DocuMark Transformer ✨")
        font = QFont(); font.setPointSize(16); font.setBold(True)
        title.setFont(font)
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setStyleSheet("padding-bottom: 10px; color: #5dade2;")

        button_style = """
            QPushButton { min-height: 40px; padding: 8px 15px; border-radius: 5px; font-size: 11pt; }
            QPushButton:disabled { background-color: #4a4a4a; color: #888; }
        """
        self.open_button = QPushButton("📂 Open Document or HTML File")
        self.copy_button = QPushButton("📋 Copy Markdown")
        self.save_button = QPushButton("💾 Save as Markdown")
        for b in (self.open_button, self.copy_button, self.save_button):
            b.setStyleSheet(button_style)

        self.drop_label = QLabel("📄 ... or drag and drop a file here.")
        self.drop_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._set_drop_hover(False)

        self.output = QTextEdit()
        self.output.setReadOnly(True)
        self.output.setPlaceholderText("Converted Markdown will appear here...")
        self.output.setStyleSheet("""
            QTextEdit { border: 1px solid #555; border-radius: 5px; padding: 10px;
                        background-color: #282828; font-size: 10pt; }
        """)

        buttons = QHBoxLayout()
        buttons.addStretch(1)
        buttons.addWidget(self.copy_button)
        buttons.addWidget(self.save_button)
        buttons.addStretch(1)

        self.status_label = QLabel()
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        layout.addWidget(title)
        layout.addWidget(self.open_button)
        layout.addWidget(self.drop_label)
        layout.addWidget(self.output, 1)
        layout.addLayout(buttons)
        layout.addWidget(self.status_label)

        self.open_button.clicked.connect(self.open_file_dialog)
        self.copy_button.clicked.connect(self.copy_markdown)
        self.save_button.clicked.connect(self.save_markdown)
        self._update_action_buttons()
        self.set_status(self.baseline_status, kind="busy")

    def _set_drop_hover(self, hover: bool):
        border, bg, fg = ("#bbb", "#3a3a3a", "#eee") if hover else ("#666", "#2a2a2a", "#aaa")
        self.drop_label.setStyleSheet(f"""
            QLabel {{ border: 2px dashed {border}; padding: 40px 20px; border-radius: 5px;
                      background-color: {bg}; font-size: 11pt; color: {fg}; }}
            QLabel:disabled {{ background-color: #333333; border-color: #444; color: #666; }}
        """)

    def _update_action_buttons(self):
        enabled = self.job_id == 0 and bool(self.output.toPlainText().strip())
        self.copy_button.setEnabled(enabled)
        self.save_button.setEnabled(enabled)

    def _set_busy(self, busy: bool):
        self.open_button.setEnabled(not busy)
        self.drop_label.setEnabled(not busy)
        self._update_action_buttons()

    def set_status(self, message: str, kind: str = "info", temporary: bool = False):
        colors = {"info": "#999", "ok": "#2ecc71", "busy": "#f39c12", "error": "#e74c3c"}
        self.status_label.setText(message)
        self.status_label.setStyleSheet(f"color: {colors[kind]}; padding-top: 5px; font-size: 9pt;")
        if temporary:
            QTimer.singleShot(3000, lambda: self.status_label.text() == message and self._restore_status())
        else:
            self._baseline = (message, kind)

    def _restore_status(self):
        self.set_status(*self._baseline)

    def show_error(self, message: str):
        QMessageBox.warning(self, "Error", message)

    # ---------------- Engine lifecycle ----------------
    def _start_engine(self):
        self.engine_thread = QThread(self)
        self.engine = Engine()
        self.engine.moveToThread(self.engine_thread)
        q = Qt.ConnectionType.QueuedConnection
        self.engine.ready.connect(self._on_engine_ready, q)
        self.engine.load_failed.connect(self._on_engine_failed, q)
        self.engine.converted.connect(self._on_converted, q)
        self.engine.failed.connect(self._on_failed, q)
        self.submit.connect(self.engine.convert, q)
        self.engine_thread.started.connect(self.engine.load)
        # Let the window paint before the import starts competing for the GIL.
        QTimer.singleShot(50, self.engine_thread.start)

    @Slot()
    def _on_engine_ready(self):
        self.engine_state = "ready"
        mark("engine ready")
        if self.job_id:
            self._set_progress_text(f"Converting '{os.path.basename(self.last_processed_file)}'...")
        else:
            self.set_status(self.READY_TEXT)

    @Slot(str, str)
    def _on_engine_failed(self, error_type: str, message: str):
        self.engine_state = "failed"
        self.engine_error = f"{error_type}: {message}"
        msg = f"Failed to load Docling.\n\nError: {self.engine_error}\n\n"
        if error_type in ("ImportError", "ModuleNotFoundError"):
            msg += "Install the dependencies with:\n  pip install -r requirements.txt"
        if self.job_id:
            self._finish_job()
        self.set_status("❌ Docling failed to load. See the error message.", kind="error")
        QMessageBox.critical(self, "Initialization Error", msg)

    def closeEvent(self, event):
        if self.progress:
            self.progress.close()
        self.engine_thread.quit()
        event.accept()

    # ---------------- Opening files ----------------
    def dragEnterEvent(self, event: QDragEnterEvent):
        urls = event.mimeData().urls() if event.mimeData().hasUrls() else []
        if not self.job_id and len(urls) == 1 and urls[0].scheme() == "file":
            event.acceptProposedAction()
            self._set_drop_hover(True)
        else:
            event.ignore()

    def dragLeaveEvent(self, event):
        self._set_drop_hover(False)
        event.accept()

    def dropEvent(self, event: QDropEvent):
        self._set_drop_hover(False)
        urls = event.mimeData().urls()
        if self.job_id or not urls:
            event.ignore()
            return
        path = urls[0].toLocalFile()
        if not path and sys.platform == "darwin":
            # Newer macOS can hand over file-reference URLs Qt cannot resolve.
            try:
                from Foundation import NSURL  # pyobjc, imported only when needed
                ns_url = NSURL.URLWithString_(urls[0].toString())
                if ns_url and ns_url.isFileURL():
                    path = str(ns_url.path())
            except ImportError:
                pass
        if not path:
            self.show_error(f"Cannot handle non-local file URL: {urls[0].toString()}")
            event.ignore()
            return
        event.acceptProposedAction()
        self.process_file(path)

    @Slot()
    def open_file_dialog(self):
        if self.job_id:
            return
        path, _ = QFileDialog.getOpenFileName(self, "Select Document or HTML File", "", OPEN_FILTERS)
        if path:
            self.process_file(path)

    # ---------------- Conversion ----------------
    def process_file(self, path: str):
        if self.job_id:
            self.show_error("Please wait for the current conversion to complete.")
            return
        if self.engine_state == "failed":
            self.show_error(f"Docling failed to load, so files cannot be converted.\n\n{self.engine_error}")
            return
        if not os.path.isfile(path):
            self.show_error(f"File not found: {path}")
            return
        if not os.access(path, os.R_OK):
            self.show_error(f"Permission denied: Cannot read file\n{path}")
            return

        self.job_id += 1
        self.cancelled = False
        self.last_processed_file = path
        name = os.path.basename(path)
        self.output.clear()
        self._set_busy(True)

        waiting = self.engine_state == "loading"
        label = (f"Loading Docling, then converting '{name}'...\n(this wait only happens right after start)"
                 if waiting else f"Converting '{name}'...")
        self.set_status(f"⏳ Converting '{name}'...", kind="busy")
        self.progress = QProgressDialog(label, "Cancel", 0, 0, self)
        self.progress.setWindowTitle("DocuMark Transformer")
        self.progress.setWindowModality(Qt.WindowModality.WindowModal)
        self.progress.setMinimumDuration(0)
        self.progress.setAutoClose(False)
        self.progress.setAutoReset(False)
        self.progress.canceled.connect(self._cancel)
        self.progress.show()

        self.submit.emit(self.job_id, path)

    def _set_progress_text(self, text: str):
        if self.progress and not self.cancelled:
            self.progress.setLabelText(text)

    @Slot()
    def _cancel(self):
        # Docling cannot be interrupted mid-document, so the result is discarded when it arrives.
        if not self.job_id:
            return
        self.cancelled = True
        if self.progress:
            self.progress.setLabelText("Cancelling... the current document has to finish first.")
            self.progress.setCancelButton(None)
            self.progress.show()   # QProgressDialog hides itself on cancel

    def _finish_job(self):
        if self.progress:
            self.progress.close()
            self.progress.deleteLater()
            self.progress = None
        self.job_id = 0
        self._set_busy(False)

    @Slot(int, str)
    def _on_converted(self, job_id: int, markdown: str):
        if job_id != self.job_id:
            return
        name = os.path.basename(self.last_processed_file)
        cancelled = self.cancelled
        self._finish_job()
        if cancelled:
            self.set_status("Conversion cancelled.")
        elif markdown.strip():
            self.output.setPlainText(markdown)
            self.set_status(f"✅ Successfully converted '{name}'!", kind="ok")
        else:
            self.set_status("❌ Conversion failed: empty result.", kind="error")
            self.show_error(f"Conversion resulted in empty content for '{name}'.")
        self._update_action_buttons()

    @Slot(int, str, str)
    def _on_failed(self, job_id: int, error_type: str, message: str):
        if job_id != self.job_id:
            return
        name = os.path.basename(self.last_processed_file)
        cancelled = self.cancelled
        self._finish_job()
        if cancelled:
            self.set_status("Conversion cancelled.")
            return
        msg = f"Failed to convert '{name}'."
        if error_type in ("OSError", "PermissionError", "FileNotFoundError"):
            if sys.platform == "win32" and "1314" in message:
                msg += ("\n\nPrivilege Error (WinError 1314).\n\nTroubleshooting:\n"
                        "- Enable Windows Developer Mode, or run as Administrator.\n"
                        "- Check permissions for the Hugging Face cache:\n"
                        f"  {os.path.expanduser('~')}\\.cache\\huggingface")
            else:
                msg += f"\n\nOS Error: {message}\n\nCheck file permissions and if the file is open elsewhere."
        elif error_type in ("ImportError", "ModuleNotFoundError"):
            msg += f" A required dependency might be missing.\n\nDetails: {message}"
        else:
            msg += f"\n\n{error_type}: {message}"
        self.set_status(f"❌ Conversion failed ({error_type}).", kind="error")
        self.show_error(msg)

    # ---------------- Copy / save ----------------
    @Slot()
    def copy_markdown(self):
        text = self.output.toPlainText()
        if text:
            QApplication.clipboard().setText(text)
            self.set_status("📋 Markdown copied to clipboard!", kind="ok", temporary=True)

    @Slot()
    def save_markdown(self):
        text = self.output.toPlainText()
        if not text:
            return
        stem = os.path.splitext(os.path.basename(self.last_processed_file or "output"))[0]
        path, chosen = QFileDialog.getSaveFileName(self, "Save Markdown As", f"{stem}.md", SAVE_FILTERS)
        if not path:
            return
        lower = path.lower()
        if chosen.startswith("Markdown") and not lower.endswith((".md", ".markdown")):
            path += ".md"
        elif chosen.startswith("Text") and not lower.endswith(".txt"):
            path += ".txt"
        elif "." not in os.path.basename(path):
            path += ".md"
        try:
            with open(path, "w", encoding="utf-8") as f:
                f.write(text)
            self.set_status(f"💾 Saved to '{os.path.basename(path)}'", kind="ok", temporary=True)
        except OSError as e:
            self.set_status("❌ File save failed.", kind="error", temporary=True)
            self.show_error(f"Could not save file: {e}\n\nCheck directory permissions and available disk space.")


def main():
    app = QApplication(sys.argv)
    try:
        import qdarkstyle
        app.setStyleSheet(qdarkstyle.load_stylesheet(qt_api="pyside6"))
    except Exception as e:  # theme is optional
        print(f"Warning: dark theme unavailable ({e}). Install with: pip install qdarkstyle")
    window = MarkdownConverterApp()
    window.show()
    QTimer.singleShot(0, lambda: mark("window shown"))
    code = app.exec()
    # Exit straight away: a conversion may still be running in the engine thread
    # (Docling cannot be interrupted), and tearing down torch adds a delay on exit.
    sys.stdout.flush()
    sys.stderr.flush()
    os._exit(code)


if __name__ == "__main__":
    main()
