import sys
import os
import platform
import time # For simulating load if needed
import traceback # For detailed error reporting
import threading # Used for threads

# --- FIX for Hugging Face Hub Symlink Error on Windows ---
if platform.system() == 'Windows':
    print("Applying Windows Hugging Face Hub symlink workaround...")
    os.environ['HF_HUB_DISABLE_SYMLINKS'] = '1'
    os.environ['HF_HUB_DISABLE_SYMLINKS_WARNING'] = '1'
# --- End Fix ---

# --- GUI Imports (Place after potential env var changes) ---
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QFileDialog, QTextEdit, QLabel, QMessageBox, QSizePolicy,
    QProgressDialog, QProgressBar, QSpacerItem # Added QSpacerItem
)
from PySide6.QtCore import Qt, Slot, QUrl, QTimer, QSize, QThread, Signal, QObject
from PySide6.QtGui import QDragEnterEvent, QDropEvent, QClipboard, QFont, QMovie
import qdarkstyle # For dark theme

# --- macOS Drag-and-Drop Workaround ---
if platform.system() == 'Darwin':
    try:
        from Foundation import NSURL
        MACOS_DRAG_DROP_WORKAROUND = True
    except ImportError:
        MACOS_DRAG_DROP_WORKAROUND = False
else:
    MACOS_DRAG_DROP_WORKAROUND = False

# --- Docling Import - Defer this to happen after GUI shows ---
# Will be imported in InitializationWorker


# ==============================================================
# Worker Thread for Initialization
# ==============================================================
class InitializationWorker(QObject):
    """
    Worker object to perform the slow import in a separate thread.
    """
    initialization_complete = Signal(object) # Emits the imported class
    initialization_error = Signal(str, str)  # Emits error type, message

    @Slot()
    def run(self):
        """Performs the import."""
        print(f"[InitThread {threading.get_ident()}] Starting initialization...")
        try:
            # --- Actual Import ---
            # Ensure you have langchain-docling installed: pip install langchain-docling
            from langchain_docling import DoclingLoader as DL_Class
            print(f"[InitThread {threading.get_ident()}] DoclingLoader imported successfully.")
            # Use QTimer to ensure signal emission happens from the event loop
            QTimer.singleShot(0, lambda: self.initialization_complete.emit(DL_Class))

        except ImportError as e:
            error_type = type(e).__name__
            error_message = f"Import Error: {e}"
            print(f"[InitThread {threading.get_ident()}] {error_message}")
            QTimer.singleShot(0, lambda: self.initialization_error.emit(error_type, str(e)))
        except Exception as e:
            error_type = type(e).__name__
            tb_str = traceback.format_exc()
            error_message = f"Unexpected Initialization Error: {e}"
            print(f"[InitThread {threading.get_ident()}] {error_message}\n{tb_str}")
            QTimer.singleShot(0, lambda: self.initialization_error.emit(error_type, str(e)))
        finally:
            print(f"[InitThread {threading.get_ident()}] Initialization run method finished.")
        # No explicit cleanup here, thread `finished` signal handles calling deleteLater


# ==============================================================
# Worker Thread for Conversion
# ==============================================================
class ConversionWorker(QObject):
    """
    Worker object to perform the document conversion in a separate thread.
    """
    conversion_complete = Signal(list, str) # Emits list of docs, original file path
    conversion_error = Signal(str, str, str) # Emits error type, message, original file path
    progress_update = Signal(str) # Reports progress stages

    def __init__(self, loader_class, file_path, parent=None):
        super().__init__(parent)
        self.DoclingLoaderClass = loader_class
        self.file_path = file_path
        self._is_running = True # Flag to signal cancellation

    @Slot()
    def run(self):
        """Performs the conversion."""
        thread_id = threading.get_ident()
        print(f"[ConvThread {thread_id}] Run method started.")

        # --- Helper function to emit signals safely via QTimer ---
        def emit_signal(signal, *args):
            # Ensure signals are emitted from the thread's event loop context
            QTimer.singleShot(0, lambda: signal.emit(*args))

        if not self._is_running:
            print(f"[ConvThread {thread_id}] Worker not started, already cancelled.")
            emit_signal(self.conversion_error, "Cancelled", "Operation cancelled before starting.", self.file_path)
            return

        base_name = os.path.basename(self.file_path)
        try:
            print(f"[ConvThread {thread_id}] Starting conversion for: {self.file_path}")
            emit_signal(self.progress_update, f"Initializing conversion for '{base_name}'...")

            # Short sleep to allow UI update, check for cancellation
            time.sleep(0.1)
            if not self._is_running: raise RuntimeError("Operation cancelled during init.")

            # Initialize the loader
            loader = self.DoclingLoaderClass(file_path=self.file_path, export_type="markdown")
            print(f"[ConvThread {thread_id}] Loader initialized.")
            emit_signal(self.progress_update, f"Loading and converting '{base_name}'...")

            # Short sleep, check for cancellation
            time.sleep(0.1)
            if not self._is_running: raise RuntimeError("Operation cancelled before load.")

            # --- THE SLOW PART ---
            print(f"[ConvThread {thread_id}] Calling loader.load()...")
            docs = loader.load()
            print(f"[ConvThread {thread_id}] loader.load() finished.")
            # ---------------------

            # Check for cancellation immediately after the blocking call
            if not self._is_running: raise RuntimeError("Operation cancelled during load.")

            # If successful, emit completion signal
            print(f"[ConvThread {thread_id}] Conversion successful for: {self.file_path}")
            emit_signal(self.conversion_complete, docs, self.file_path)

        except RuntimeError as e:
             # Handle cancellation specifically
             if "cancelled" in str(e).lower():
                 print(f"[ConvThread {thread_id}] Conversion cancelled.")
                 emit_signal(self.conversion_error, "Cancelled", "Operation was cancelled by user.", self.file_path)
             else:
                 # Handle other RuntimeErrors
                 error_type = type(e).__name__
                 tb_str = traceback.format_exc()
                 print(f"[ConvThread {thread_id}] Unexpected RuntimeError: {e}\n{tb_str}")
                 emit_signal(self.conversion_error, error_type, str(e), self.file_path)

        except OSError as e:
            # Handle OS-level errors (file not found, permissions)
            error_type = type(e).__name__
            error_message = f"OS Error during conversion: {e}"
            print(f"[ConvThread {thread_id}] {error_message}")
            # Check if cancellation happened concurrently
            final_error_type = "Cancelled" if not self._is_running else error_type
            final_error_msg = "Operation cancelled after OS error." if not self._is_running else str(e)
            emit_signal(self.conversion_error, final_error_type, final_error_msg, self.file_path)
        except ImportError as e:
            # Handle missing dependencies if import happens here (less likely with init thread)
            error_type = type(e).__name__
            error_message = f"Import Error during conversion: {e}"
            print(f"[ConvThread {thread_id}] {error_message}")
            final_error_type = "Cancelled" if not self._is_running else error_type
            final_error_msg = "Operation cancelled after Import error." if not self._is_running else str(e)
            emit_signal(self.conversion_error, final_error_type, final_error_msg, self.file_path)
        except Exception as e:
            # Catch any other unexpected exceptions
            error_type = type(e).__name__
            error_message = f"Unexpected error during conversion: {e}"
            tb_str = traceback.format_exc()
            print(f"[ConvThread {thread_id}] {error_message}\n{tb_str}")
            final_error_type = "Cancelled" if not self._is_running else error_type
            final_error_msg = f"Operation cancelled after {error_type}." if not self._is_running else str(e)
            emit_signal(self.conversion_error, final_error_type, final_error_msg, self.file_path)
        finally:
             # This block executes whether an exception occurred or not
             print(f"[ConvThread {thread_id}] Conversion run method finished.")
        # The thread's event loop should continue running until quit() is called.

    @Slot()
    def stop(self):
        """Signals the worker to stop (best effort)."""
        print(f"[ConvThread {threading.get_ident()}] Received stop signal.")
        self._is_running = False


# ==============================================================
# Main Application Window
# ==============================================================
class MarkdownConverterApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setGeometry(100, 100, 800, 650)
        self.setWindowTitle("DocuMark Transformer (Initializing...)")

        # --- State Attributes ---
        self.original_status_text = "Initializing..."
        self.ready_status_text = "Ready. Select a file or drag it here."
        self.last_processed_file = None
        self.DoclingLoaderClass = None # Will hold the imported class
        self.is_processing = False    # Flag to prevent concurrent operations

        # --- Threading Attributes ---
        self.init_thread = None
        self.init_worker = None
        self.conversion_thread = None
        self.conversion_worker = None # Renamed from self.worker for clarity

        # --- UI Elements ---
        self.progress_dialog = None # Reference to the progress dialog

        # --- Central Widget & Layout ---
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.main_layout = QVBoxLayout(self.central_widget)
        self.main_layout.setContentsMargins(20, 20, 20, 15)
        self.main_layout.setSpacing(15)
        self.main_layout.setAlignment(Qt.AlignmentFlag.AlignCenter) # Initially center loading

        # --- Loading State Widgets ---
        self.loading_label = QLabel("🚀 Initializing DocuMark Transformer...")
        loading_font = QFont()
        loading_font.setPointSize(14)
        self.loading_label.setFont(loading_font)
        self.loading_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.loading_label.setStyleSheet("color: #cccccc; padding-bottom: 10px;")

        self.init_progress_bar = QProgressBar(self)
        self.init_progress_bar.setRange(0, 0) # Indeterminate
        self.init_progress_bar.setTextVisible(False)
        self.init_progress_bar.setMaximumHeight(10)
        self.init_progress_bar.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)

        self.main_layout.addWidget(self.loading_label)
        self.main_layout.addWidget(self.init_progress_bar)
        self.main_layout.addStretch(1) # Keep loading UI centered vertically

        # --- Main UI Widgets (Created but initially hidden) ---
        self.title_label = QLabel("✨ DocuMark Transformer ✨")
        title_font = QFont()
        title_font.setPointSize(16); title_font.setBold(True)
        self.title_label.setFont(title_font)
        self.title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.title_label.setStyleSheet("padding-bottom: 10px; color: #5dade2;")
        self.title_label.setVisible(False)

        self.open_button = QPushButton("📂 Open Document or HTML File")
        self.open_button.setVisible(False)

        self.drop_label = QLabel("📄 ... or drag and drop a file here.")
        self.drop_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.drop_label.setVisible(False)

        self.markdown_output = QTextEdit()
        self.markdown_output.setReadOnly(True)
        self.markdown_output.setPlaceholderText("Converted Markdown will appear here...")
        self.markdown_output.setVisible(False)

        # Use a container widget for the button layout for easier hiding/showing
        self.button_layout_widget = QWidget()
        self.button_layout_widget.setVisible(False)
        self.button_layout = QHBoxLayout(self.button_layout_widget)
        self.copy_button = QPushButton("📋 Copy Markdown")
        self.save_button = QPushButton("💾 Save as Markdown")
        self.copy_button.setEnabled(False)
        self.save_button.setEnabled(False)
        self.button_layout.addStretch(1)
        self.button_layout.addWidget(self.copy_button)
        self.button_layout.addWidget(self.save_button)
        self.button_layout.addStretch(1)
        self.button_layout.setContentsMargins(0, 0, 0, 0) # No extra margins for the layout itself

        self.status_label = QLabel(self.original_status_text)
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        # Status label is added later

        # --- Apply Styles to hidden widgets ---
        button_min_height = 40
        button_padding = "8px 15px"; border_radius = "5px"
        common_button_style = f"""
            QPushButton {{
                min-height: {button_min_height}px;
                padding: {button_padding};
                border-radius: {border_radius};
                font-size: 11pt;
                /* Add default background/color if needed */
            }}
            QPushButton:disabled {{
                background-color: #4a4a4a; /* Example disabled style */
                color: #888;
            }}
        """
        self.open_button.setStyleSheet(common_button_style)
        self.copy_button.setStyleSheet(common_button_style)
        self.save_button.setStyleSheet(common_button_style)

        icon_size = QSize(18, 18); self.open_button.setIconSize(icon_size); self.copy_button.setIconSize(icon_size); self.save_button.setIconSize(icon_size)

        # Define base and hover styles for drop label
        self.base_drop_style = f"""
            QLabel {{
                border: 2px dashed #666;
                padding: 40px 20px;
                border-radius: {border_radius};
                background-color: #2a2a2a;
                font-size: 11pt;
                color: #aaa;
            }}
            QLabel:disabled {{
                background-color: #333333;
                border-color: #444;
                color: #666;
            }}
        """
        self.hover_drop_style = f"""
            QLabel {{
                border: 2px dashed #bbb; /* Brighter border on hover */
                padding: 40px 20px;
                border-radius: {border_radius};
                background-color: #3a3a3a; /* Slightly lighter background */
                font-size: 11pt;
                color: #eee;
            }}
        """
        self.drop_label.setStyleSheet(self.base_drop_style)

        self.markdown_output.setStyleSheet(f"""
            QTextEdit {{
                border: 1px solid #555;
                border-radius: {border_radius};
                padding: 10px;
                background-color: #282828;
                font-size: 10pt;
            }}
            QTextEdit:disabled {{
                background-color: #333333;
                color: #888;
            }}
        """)

        self.status_label.setStyleSheet("QLabel { color: #999; padding-top: 5px; font-size: 9pt; }")

        # --- Add Main UI Widgets to Layout (Initially Hidden) ---
        # Insert before the stretch item and the status label which will be added last
        insert_index = self.main_layout.count() - 1 # Index before the stretch item
        self.main_layout.insertWidget(insert_index, self.title_label)
        self.main_layout.insertWidget(insert_index + 1, self.open_button)
        self.main_layout.insertWidget(insert_index + 2, self.drop_label)
        self.main_layout.insertWidget(insert_index + 3, self.markdown_output, 1) # Stretch factor 1
        self.main_layout.insertWidget(insert_index + 4, self.button_layout_widget)
        self.main_layout.addWidget(self.status_label) # Add status label at the very end

        # --- Enable Drag and Drop ---
        self.setAcceptDrops(True) # Enable drops on the main window

        # --- Connections ---
        self.open_button.clicked.connect(self.open_file_dialog)
        self.copy_button.clicked.connect(self.copy_markdown_to_clipboard)
        self.save_button.clicked.connect(self.save_markdown_to_file)
        self.markdown_output.textChanged.connect(self.update_action_buttons_state)

        # --- Start Initialization Thread ---
        # Use QTimer.singleShot to ensure the main event loop is running first
        QTimer.singleShot(50, self.start_initialization_thread)

    # --- Initialization Handling ---
    @Slot()
    def start_initialization_thread(self):
        """Sets up and starts the initialization worker thread."""
        if self.init_thread: # Prevent starting if already running
             print("Initialization thread already exists.")
             return
        print("Starting initialization thread setup...")
        self.init_thread = QThread(self)
        self.init_worker = InitializationWorker()
        self.init_worker.moveToThread(self.init_thread)

        # Connect signals from worker to main thread slots
        self.init_worker.initialization_complete.connect(self.handle_initialization_complete)
        self.init_worker.initialization_error.connect(self.handle_initialization_error)

        # Connect thread signals
        self.init_thread.started.connect(self.init_worker.run)
        # Clean up worker and thread when the thread finishes
        self.init_thread.finished.connect(self.init_worker.deleteLater)
        self.init_thread.finished.connect(self.init_thread.deleteLater)
        # Also clear our references when the thread is finished
        self.init_thread.finished.connect(self._clear_init_thread_references)

        print("Starting initialization thread...")
        self.init_thread.start()

    @Slot(object)
    def handle_initialization_complete(self, imported_class):
        """Handles successful initialization."""
        print("Handling initialization complete.")
        self.DoclingLoaderClass = imported_class
        self.show_main_ui()
        # References cleared via _clear_init_thread_references slot

    @Slot(str, str)
    def handle_initialization_error(self, error_type, error_message):
        """Handles errors during initialization."""
        print(f"Handling initialization Error ({error_type}): {error_message}")
        msg = f"Failed to initialize application.\n\nError: {error_type}: {error_message}\n\n"
        if error_type == "ImportError":
             msg += "Please ensure 'langchain-docling' and its dependencies are installed correctly.\n(e.g., pip install langchain-docling)"
        else:
             msg += "See console output for more details."

        self.show_initialization_error(msg)
        # References cleared via _clear_init_thread_references slot

    @Slot()
    def _clear_init_thread_references(self):
        """Slot connected to init_thread.finished to clear references."""
        print("Clearing init thread/worker references.")
        self.init_worker = None
        self.init_thread = None

    @Slot()
    def show_main_ui(self):
        """Hides loading indicators and shows the main application UI."""
        print("Showing main UI.")
        # Hide loading widgets
        self.loading_label.setVisible(False)
        self.init_progress_bar.setVisible(False)
        # Remove them from layout to reclaim space (optional but cleaner)
        self.main_layout.removeWidget(self.loading_label)
        self.main_layout.removeWidget(self.init_progress_bar)

        # Remove the stretch item added for centering the loading UI
        items_to_remove = []
        for i in range(self.main_layout.count()):
            item = self.main_layout.itemAt(i)
            # Check if it's a QSpacerItem and it provides vertical stretch
            if isinstance(item, QSpacerItem) and item.expandingDirections() == Qt.Orientation.Vertical:
                items_to_remove.append(item)
                print("Found potential loading stretch item.")
        # Remove found stretch items
        for item in items_to_remove:
            self.main_layout.removeItem(item)
            print("Removed loading stretch item.")

        # Change layout alignment now that content is present
        self.main_layout.setAlignment(Qt.AlignmentFlag.AlignTop) # Align main content to top

        # Show the main widgets
        self.title_label.setVisible(True)
        self.open_button.setVisible(True)
        self.drop_label.setVisible(True)
        self.markdown_output.setVisible(True)
        self.button_layout_widget.setVisible(True)
        self.status_label.setVisible(True) # Ensure status label is visible

        # Update window state
        self.setWindowTitle("DocuMark Transformer - Convert Documents to Markdown")
        self.original_status_text = self.ready_status_text # Update baseline status
        self.set_status(self.original_status_text)
        self.setEnabled(True) # Enable the main window
        self.central_widget.setEnabled(True) # Enable the central widget


    @Slot()
    def show_initialization_error(self, message):
        """Displays a critical error message if initialization fails."""
        print(f"Displaying initialization error: {message}")
        # Hide loading indicators even on error
        self.loading_label.setVisible(False)
        self.init_progress_bar.setVisible(False)
        # Optionally remove them
        self.main_layout.removeWidget(self.loading_label)
        self.main_layout.removeWidget(self.init_progress_bar)

        # Update status to indicate failure
        self.set_status("❌ Initialization Failed!")
        self.status_label.setStyleSheet("color: #e74c3c; padding-top: 5px; font-size: 9pt;") # Error color

        # Show a critical message box
        QMessageBox.critical(self, "Initialization Error", message)

        # Disable the main UI as the app can't function
        self.central_widget.setEnabled(False)
        # Optionally disable the whole window, but disabling central widget is usually enough
        # self.setEnabled(False)


    # --- Event Handlers (Drag/Drop, Open) ---
    def dragEnterEvent(self, event: QDragEnterEvent):
        """Handles drag enter events."""
        # Accept drops only if initialized, not processing, and data is a single local file URL
        if self.DoclingLoaderClass and not self.is_processing:
            mime_data = event.mimeData()
            if mime_data.hasUrls():
                urls = mime_data.urls()
                if len(urls) == 1:
                    url = urls[0]
                    # Check if it's a local file (works for most OS) or use macOS workaround
                    is_local = url.isLocalFile()
                    is_macos_file = MACOS_DRAG_DROP_WORKAROUND and url.scheme() == 'file'
                    if is_local or is_macos_file:
                        event.acceptProposedAction()
                        self.drop_label.setStyleSheet(self.hover_drop_style) # Provide visual feedback
                        return # Accepted
        event.ignore() # Ignore in all other cases

    def dragLeaveEvent(self, event):
        """Handles drag leave events."""
        # Reset drop label style if it's enabled
        if self.drop_label.isEnabled():
            self.drop_label.setStyleSheet(self.base_drop_style)
        event.accept()

    def dropEvent(self, event: QDropEvent):
        """Handles drop events."""
        # Reset drop label style first
        if self.drop_label.isEnabled():
             self.drop_label.setStyleSheet(self.base_drop_style)

        # Check conditions again (safety)
        if not (self.DoclingLoaderClass and not self.is_processing):
            event.ignore()
            return

        mime_data = event.mimeData()
        if mime_data.hasUrls():
            urls = mime_data.urls()
            if urls: # Should be exactly one due to dragEnterEvent logic
                url = urls[0]
                file_path = ""
                if url.isLocalFile():
                    file_path = url.toLocalFile()
                elif MACOS_DRAG_DROP_WORKAROUND and url.scheme() == 'file':
                    # macOS specific handling using Foundation
                    try:
                        # Create an NSURL object from the string representation of the QUrl
                        ns_url = NSURL.URLWithString_(url.toString())
                        # Check if it's a file URL and get the path
                        if ns_url and ns_url.isFileURL():
                            file_path = str(ns_url.path()) # Get path string
                        else:
                            self.show_error(f"Cannot resolve non-file URL on macOS: {url.toString()}")
                            event.ignore(); return
                    except Exception as e:
                        self.show_error(f"Error resolving macOS path: {e}\nURL: {url.toString()}")
                        event.ignore(); return
                else:
                    # Handle cases where it's not a local file (e.g., http URL)
                    self.show_error(f"Cannot handle non-local file URL: {url.toString()}")
                    event.ignore(); return

                # If we got a valid file path, process it
                if file_path:
                    print(f"File dropped: {file_path}")
                    event.acceptProposedAction()
                    self.process_file(file_path)
                    return # Handled

        event.ignore() # Ignore if no valid URL found

    @Slot()
    def open_file_dialog(self):
        """Opens a file dialog to select a file."""
        if not self.DoclingLoaderClass or self.is_processing:
            return # Don't open if not ready or busy

        # Define supported file types
        supported_filters = (
            "All Supported Files (*.pdf *.docx *.pptx *.html *.htm);;"
            "PDF Files (*.pdf);;"
            "Word Documents (*.docx);;"
            "PowerPoint Files (*.pptx);;"
            "HTML Files (*.html *.htm);;"
            "All Files (*)"
        )

        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Document or HTML File",
            "", # Start directory (empty means default/last used)
            supported_filters
        )

        if file_path: # Proceed if a file was selected
            print(f"File selected via dialog: {file_path}")
            self.process_file(file_path)


    # --- Conversion Process ---
    def process_file(self, file_path: str):
        """Initiates the file processing in a separate thread with progress dialog."""
        # Pre-checks
        if not self.DoclingLoaderClass:
            self.show_error("Application is not fully initialized.")
            return
        if self.is_processing:
            self.show_error("Please wait for the current conversion to complete.")
            return
        if not os.path.exists(file_path):
            self.show_error(f"File not found: {file_path}")
            self.reset_status("File access error.")
            return
        if not os.access(file_path, os.R_OK):
            self.show_error(f"Permission denied: Cannot read file\n{file_path}")
            self.reset_status("File permission error.")
            return

        # --- Start Processing State ---
        self.is_processing = True
        self.last_processed_file = file_path
        base_name = os.path.basename(file_path)

        print(f"Starting conversion process for: {file_path}")
        self.set_status(f"⏳ Preparing conversion for '{base_name}'...", is_processing=True)
        self.markdown_output.clear() # Clear previous output
        self.markdown_output.setPlaceholderText(f"Starting conversion for '{base_name}'...")
        self.update_action_buttons_state() # Disables copy/save
        self.set_ui_enabled(False) # Disable open/drop during processing

        # --- Setup and Show Progress Dialog ---
        # Ensure any previous dialog reference is cleared (safety)
        if self.progress_dialog:
             print("Warning: Previous progress dialog reference exists. Closing it.")
             self.progress_dialog.close()
             self.progress_dialog = None

        self.progress_dialog = QProgressDialog(
            f"Converting '{base_name}'...", # Label text
            "Cancel",                      # Cancel button text
            0,                             # Minimum value (0 for indeterminate)
            0,                             # Maximum value (0 for indeterminate)
            self                           # Parent widget
        )
        self.progress_dialog.setWindowModality(Qt.WindowModality.WindowModal) # Block main window
        self.progress_dialog.setMinimumDuration(0) # Show immediately
        self.progress_dialog.setAutoClose(False)   # Don't close automatically
        self.progress_dialog.setAutoReset(False)   # Don't reset automatically
        self.progress_dialog.canceled.connect(self.cancel_conversion) # Connect cancel signal
        print("Showing progress dialog.")
        self.progress_dialog.show()
        QApplication.processEvents() # Crucial: Ensure dialog is shown *before* thread starts blocking

        # --- Setup and Start Thread ---
        # Ensure previous thread references are cleared (safety)
        if self.conversion_thread:
             print("Warning: Previous conversion thread reference exists.")
             # Attempt cleanup if needed, though handle_conversion_finished should do this
             if self.conversion_thread.isRunning():
                 self.conversion_thread.quit()
                 self.conversion_thread.wait(100) # Brief wait
             self.conversion_thread = None
             self.conversion_worker = None

        self.conversion_thread = QThread(self)
        self.conversion_worker = ConversionWorker(self.DoclingLoaderClass, file_path)
        self.conversion_worker.moveToThread(self.conversion_thread)

        # Connect signals from worker to slots in this main thread
        self.conversion_worker.conversion_complete.connect(self.handle_conversion_complete)
        self.conversion_worker.conversion_error.connect(self.handle_conversion_error)
        # Ensure progress dialog is updated (check if it exists)
        self.conversion_worker.progress_update.connect(
            lambda msg: self.progress_dialog.setLabelText(msg) if self.progress_dialog else None
        )

        # Connect thread management signals
        self.conversion_thread.started.connect(self.conversion_worker.run)
        # *** CRITICAL: Connect finished signal to the central cleanup slot ***
        self.conversion_thread.finished.connect(self.handle_conversion_finished)

        print("Starting conversion thread...")
        self.conversion_thread.start()


    @Slot()
    def cancel_conversion(self):
        """Slot called when the progress dialog's Cancel button is clicked."""
        print("Cancel button clicked.")
        # Check if worker exists and we are actually processing
        if self.conversion_worker and self.is_processing:
            print("Signalling worker thread to stop...")
            self.conversion_worker.stop() # Tell the worker to stop
            if self.progress_dialog:
                # Update dialog to show cancellation is in progress
                self.progress_dialog.setLabelText("Cancelling conversion, please wait...")
                self.progress_dialog.setEnabled(False) # Disable cancel button now
        else:
            print("Cancel clicked but no active worker or not processing.")
            # If cancel is clicked but processing already finished somehow,
            # handle_conversion_finished should still clean up the dialog.


    @Slot(list, str)
    def handle_conversion_complete(self, docs, original_file_path):
        """Handles successful completion signal from the worker thread."""
        print(f"Handling conversion complete signal for: {original_file_path}")

        # --- State Check ---
        # Ensure this signal corresponds to the currently processed file and state
        if original_file_path != self.last_processed_file or not self.is_processing:
            print("Warning: Received completion signal for outdated/unexpected file or state. Ignoring.")
            # Even if ignored, handle_conversion_finished will still run via thread.finished
            # Explicitly quit the thread here just in case.
            if self.conversion_thread:
                print("Quitting thread due to outdated completion signal.")
                self.conversion_thread.quit()
            return

        # --- Process Results ---
        base_name = os.path.basename(original_file_path)
        if docs:
            # Extract page content, filtering out potential None values
            page_contents = [doc.page_content for doc in docs if hasattr(doc, 'page_content') and doc.page_content]
            if page_contents:
                full_markdown = "\n\n".join(page_contents)
                self.markdown_output.setText(full_markdown)
                self.set_status(f"✅ Successfully converted '{base_name}'!", is_success=True)
            else:
                # Handle case where docs exist but have no content
                self.markdown_output.setPlaceholderText(f"Conversion of '{base_name}' resulted in empty content.")
                self.show_error(f"Conversion resulted in empty content for '{base_name}'.")
                self.reset_status("Conversion failed: Empty result.")
        else:
            # Handle case where the loader returned an empty list or None
            self.markdown_output.setPlaceholderText(f"No processable documents found in '{base_name}'.")
            self.show_error(f"Docling returned no processable documents for '{base_name}'.")
            self.reset_status("Conversion failed: No documents.")

        # --- Explicitly tell the thread to quit ---
        # This encourages the finished signal to be emitted.
        if self.conversion_thread:
            print("Explicitly quitting conversion thread after handling completion.")
            self.conversion_thread.quit()
        # --- IMPORTANT ---
        # UI re-enabling and final cleanup STILL happens in handle_conversion_finished
        # which is triggered by the thread's finished() signal.


    @Slot(str, str, str)
    def handle_conversion_error(self, error_type, error_message_str, original_file_path):
        """Handles error signal from the worker thread."""
        print(f"Handling conversion error signal ({error_type}) for: {original_file_path}")

        # --- State Check ---
        if original_file_path != self.last_processed_file or not self.is_processing:
             print("Warning: Received error signal for outdated/unexpected file or state. Ignoring.")
             # Explicitly quit the thread here just in case.
             if self.conversion_thread:
                 print("Quitting thread due to outdated error signal.")
                 self.conversion_thread.quit()
             return

        # --- Process Error ---
        base_name = os.path.basename(original_file_path)
        self.markdown_output.clear() # Clear any partial output or placeholder

        # Construct user-friendly messages based on error type
        user_msg = f"Failed to convert '{base_name}'."
        status_msg = f"❌ Conversion failed ({error_type})." # Default status

        if error_type == "Cancelled":
            user_msg = f"Conversion for '{base_name}' was cancelled."
            status_msg = "🤷‍♀️ Conversion Cancelled."
            self.markdown_output.setPlaceholderText("Conversion was cancelled.")
        elif error_type == "OSError":
            # Specific check for Windows privilege error (common with symlinks/cache)
            if platform.system() == 'Windows' and "1314" in error_message_str:
                 user_msg += (f"\n\nPrivilege Error (WinError 1314).\n\n"
                              f"Troubleshooting:\n- Try running as Administrator.\n"
                              f"- Check permissions for Hugging Face cache:\n"
                              f"  C:\\Users\\{os.getlogin()}\\.cache\\huggingface")
            else:
                 user_msg += f"\n\nOS Error: {error_message_str}\n\nCheck file permissions and if the file is open elsewhere."
            status_msg = "❌ Conversion failed (OS Error)."
            self.markdown_output.setPlaceholderText("Conversion failed (OS Error). See message.")
        elif error_type == "ImportError":
            user_msg += f" A required dependency might be missing.\n\nDetails: {error_message_str}"
            status_msg = "❌ Conversion failed (Missing Dependency)."
            self.markdown_output.setPlaceholderText("Conversion failed (Missing Dependency). See message.")
        else: # General Exception
            user_msg += f"\n\nUnexpected Error: {error_type}\n\nDetails: {error_message_str}\n\nSee console for detailed traceback."
            self.markdown_output.setPlaceholderText(f"Conversion failed ({error_type}). See message.")

        # Show error message box only if it wasn't a user cancellation
        if error_type != "Cancelled":
            self.show_error(user_msg) # show_error uses QTimer.singleShot for safety

        # Update the status bar (reset_status uses QTimer.singleShot)
        self.reset_status(status_msg)

        # --- Explicitly tell the thread to quit ---
        # This encourages the finished signal to be emitted.
        if self.conversion_thread:
            print("Explicitly quitting conversion thread after handling error.")
            self.conversion_thread.quit()
        # --- IMPORTANT ---
        # UI re-enabling and final cleanup STILL happens in handle_conversion_finished


    @Slot()
    def handle_conversion_finished(self):
        """Central cleanup called when the conversion thread finishes (success, error, or cancel)."""
        # This slot is connected to the thread's finished() signal.
        print(">>> handle_conversion_finished STARTING <<<") # Diagnostic print

        # --- Close and clean up progress dialog FIRST ---
        print("Checking progress dialog...")
        if self.progress_dialog:
            print("Closing progress dialog...")
            # It's important to close it before potentially blocking operations
            self.progress_dialog.close()
            # Schedule for deletion if it makes sense, but setting to None is key
            # self.progress_dialog.deleteLater() # Usually not needed if closed
            self.progress_dialog = None # Crucial: Release reference immediately
            print("Progress dialog closed and reference cleared.")
        else:
            print("Progress dialog reference was already None.")

        # --- Defer rest of UI updates slightly using QTimer ---
        # This allows the event loop to process the dialog closing first.
        QTimer.singleShot(0, self._finalize_conversion_ui)

        # --- Clean up thread and worker objects ---
        # This can happen immediately after signaling the UI update timer.
        # The objects will be deleted when control returns to the event loop.
        print("Scheduling thread and worker deletion...")
        if self.conversion_worker:
            self.conversion_worker.deleteLater()
            self.conversion_worker = None
        if self.conversion_thread:
            # Thread should already be quit by handle_complete/error,
            # but quit() here is safe. deleteLater schedules deletion.
            self.conversion_thread.quit()
            self.conversion_thread.deleteLater()
            self.conversion_thread = None
        print("Thread and worker scheduled for deletion.")

        print(">>> handle_conversion_finished COMPLETED (UI update deferred) <<<")

    @Slot()
    def _finalize_conversion_ui(self):
        """Helper slot called by QTimer to finalize UI updates after conversion."""
        print(">>> _finalize_conversion_ui STARTING <<<")

        # --- Reset processing flag ---
        # Needs to happen before enabling UI/buttons that depend on it
        print(f"Setting is_processing from {self.is_processing} to False")
        self.is_processing = False

        # --- Re-enable main UI elements ---
        print("Calling set_ui_enabled(True)...")
        self.set_ui_enabled(True)
        print("set_ui_enabled(True) called.")

        # --- Reset placeholder text if output is still empty ---
        # This might happen on error or if conversion yielded nothing
        if not self.markdown_output.toPlainText().strip():
            self.markdown_output.setPlaceholderText("Converted Markdown will appear here...")
            print("Placeholder text reset (if needed).")

        # --- Update button states ---
        # Relies on is_processing being False now
        print("Calling update_action_buttons_state()...")
        self.update_action_buttons_state()
        print("update_action_buttons_state() called.")

        # --- Check and reset status if needed ---
        # If the status still shows a processing state, reset it to the default ready state.
        current_status = self.status_label.text()
        if current_status.startswith("⏳") or "Preparing" in current_status or "Initializing" in current_status or "Converting" in current_status:
            print(f"Resetting status label from '{current_status}' as it shows processing state.")
            # Use reset_status which handles setting the correct baseline text
            self.reset_status()
        else:
             print(f"Status label ('{current_status}') doesn't show processing, not resetting.")

        print(">>> _finalize_conversion_ui COMPLETED <<<")


    # --- Action Button Slots (Copy, Save) ---
    @Slot()
    def copy_markdown_to_clipboard(self):
        """Copies the content of the markdown output to the clipboard."""
        markdown_text = self.markdown_output.toPlainText()
        if markdown_text:
            try:
                QApplication.clipboard().setText(markdown_text)
                self.set_status("📋 Markdown copied to clipboard!", is_success=True, temporary=True)
            except Exception as e:
                # Handle potential clipboard errors (rare)
                self.show_error(f"Could not copy to clipboard: {e}")
                self.set_status("❌ Clipboard copy failed.", temporary=True)
        else:
            # No text to copy
            self.set_status("🤷‍♀️ Nothing to copy.", temporary=True)

    @Slot()
    def save_markdown_to_file(self):
        """Opens a save file dialog to save the markdown content."""
        markdown_text = self.markdown_output.toPlainText()
        if not markdown_text:
            self.set_status("🤷‍♀️ Nothing to save.", temporary=True)
            return

        # Suggest a filename based on the last processed file
        default_filename = "output.md"
        if self.last_processed_file:
            base = os.path.basename(self.last_processed_file)
            name_without_ext = os.path.splitext(base)[0]
            default_filename = f"{name_without_ext}.md"

        # File type filters for the save dialog
        save_filters = "Markdown Files (*.md);;Text Files (*.txt);;All Files (*)"

        file_path, selected_filter = QFileDialog.getSaveFileName(
            self,
            "Save Markdown As",
            default_filename, # Suggested filename/path
            save_filters
        )

        if file_path: # Proceed if a path was chosen
            # Automatically add extension if missing based on filter (optional but helpful)
            try:
                # Check lower case extension to be safe
                file_lower = file_path.lower()
                if selected_filter == "Markdown Files (*.md)" and not file_lower.endswith((".md", ".markdown")):
                    file_path += ".md"
                elif selected_filter == "Text Files (*.txt)" and not file_lower.endswith(".txt"):
                    file_path += ".txt"
                # Basic check if *any* extension is missing when "All Files" might be used
                elif '.' not in os.path.basename(file_path):
                     file_path += ".md" # Default to .md if none provided

                # Write the file using UTF-8 encoding
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(markdown_text)

                saved_filename = os.path.basename(file_path)
                self.set_status(f"💾 Saved to '{saved_filename}'", is_success=True, temporary=True)

            except OSError as e:
                # Handle file system errors (permissions, disk full, etc.)
                self.show_error(f"Could not save file: {e}\n\nCheck directory permissions and available disk space.")
                self.set_status("❌ File save failed (OS Error).", temporary=True)
            except Exception as e:
                # Handle other unexpected errors during save
                tb_str = traceback.format_exc()
                print(f"Save Error: {e}\n{tb_str}")
                self.show_error(f"Could not save file: {type(e).__name__}: {e}")
                self.set_status("❌ File save failed.", temporary=True)


    # --- UI State Management ---
    @Slot()
    def update_action_buttons_state(self):
        """Enables/disables Copy and Save buttons based on text content and processing state."""
        has_text = bool(self.markdown_output.toPlainText().strip())
        # Buttons should be enabled only if there's text AND not currently processing AND initialized
        can_interact = has_text and not self.is_processing and self.DoclingLoaderClass is not None

        self.copy_button.setEnabled(can_interact)
        self.save_button.setEnabled(can_interact)
        # print(f"Action buttons updated: Enabled={can_interact} (HasText={has_text}, NotProcessing={not self.is_processing}, Initialized={self.DoclingLoaderClass is not None})")


    def set_ui_enabled(self, enabled: bool):
        """Enables or disables primary interaction widgets."""
        self.open_button.setEnabled(enabled)
        self.drop_label.setEnabled(enabled)
        # Also update the visual state of the drop label via property for styling
        self.drop_label.setProperty("enabled", enabled)
        # Re-polish the style to apply :disabled state changes
        self.drop_label.style().unpolish(self.drop_label)
        self.drop_label.style().polish(self.drop_label)
        # The markdown output and copy/save buttons are handled by update_action_buttons_state

    def set_status(self, message: str, is_success=False, is_processing=False, temporary=False):
        """Updates the status bar label with appropriate styling."""
        # Update the baseline status text only if the new message is a non-temporary, non-processing final state
        if not is_processing and not temporary:
             # Consider success, ready, cancelled, or failed messages as potential baseline states
             is_final_state = is_success or "Ready" in message or "Select a file" in message or "Cancelled" in message or "failed" in message.lower()
             if is_final_state:
                 self.original_status_text = message # Store this as the text to return to

        # Set the text
        self.status_label.setText(message)

        # Set the style based on the type of message
        base_style = "color: #999; padding-top: 5px; font-size: 9pt;" # Default grey
        style = base_style
        if is_success:
            style = "color: #2ecc71; padding-top: 5px; font-size: 9pt;" # Green for success
        elif is_processing:
            style = "color: #f39c12; padding-top: 5px; font-size: 9pt;" # Orange for processing
        elif "failed" in message.lower() or "Error" in message or "❌" in message:
            style = "color: #e74c3c; padding-top: 5px; font-size: 9pt;" # Red for error/failure

        self.status_label.setStyleSheet(style)

        # If temporary, set a timer to revert to the stored original status
        if temporary:
            # Capture the current baseline status *before* the timer lambda is created
            current_baseline = self.original_status_text
            # Use a timer to reset after 3 seconds, but only if the status hasn't changed again
            QTimer.singleShot(3000, lambda: self.reset_status(current_baseline) if self.status_label.text() == message else None)

    def reset_status(self, base_message=None):
        """Resets the status label to the stored baseline or a provided message."""
        # Determine the message to reset to
        if base_message is None:
            # Use the appropriate ready/initializing text based on current app state
            base_message = self.ready_status_text if self.DoclingLoaderClass else "Initializing..."

        # Store the new baseline if one was provided explicitly
        if base_message is not None:
             self.original_status_text = base_message

        # Use QTimer.singleShot(0, ...) to ensure the update happens in the event loop,
        # preventing potential conflicts if called rapidly.
        QTimer.singleShot(0, lambda: self.set_status(self.original_status_text))


    def show_error(self, message: str):
        """Shows a warning message box, safely using QTimer."""
        # Check if the main window/widget is enabled before showing a potentially blocking dialog
        if self.isEnabled() and self.central_widget.isEnabled():
            # Use QTimer.singleShot to ensure the message box is shown from the main event loop
            # This prevents issues if show_error is called from a non-GUI context unexpectedly
            QTimer.singleShot(0, lambda: QMessageBox.warning(self, "Error", message))
        else:
            # If the UI is disabled (e.g., during init failure), just print to console
            print(f"Suppressed Error Popup (Window/UI Disabled): {message}")


    # --- Window Close Event ---
    def closeEvent(self, event):
        """Handles the window close event."""
        print("Close event triggered.")

        # --- Attempt to gracefully stop threads ---
        # Stop Initialization Thread if running
        if self.init_thread and self.init_thread.isRunning():
            print("Stopping initialization worker/thread...")
            # Request the thread's event loop to exit
            self.init_thread.quit()
            # Optionally wait briefly for it to finish
            # self.init_thread.wait(100)

        # Stop Conversion Thread if running
        if self.conversion_thread and self.conversion_thread.isRunning():
            print("Stopping conversion worker/thread...")
            # Signal the worker to stop its operation (best effort)
            if self.conversion_worker:
                self.conversion_worker.stop()
            # Request the thread's event loop to exit
            self.conversion_thread.quit()
            # Optionally wait
            # self.conversion_thread.wait(100)

        # --- Close Progress Dialog if open ---
        if self.progress_dialog:
            print("Closing progress dialog on exit.")
            self.progress_dialog.close()

        # Accept the close event to allow the window to close
        print("Accepting close event.")
        event.accept()


if __name__ == "__main__":
    # Ensure QApplication is created first
    app = QApplication(sys.argv)

    # Apply dark theme (optional, requires qdarkstyle)
    try:
        # Specify the Qt API being used (PySide6)
        stylesheet = qdarkstyle.load_stylesheet(qt_api='pyside6')
        app.setStyleSheet(stylesheet)
        print("Applied qdarkstyle stylesheet.")
    except ImportError:
         print("Warning: qdarkstyle not found. Install with: pip install qdarkstyle")
    except Exception as e:
        print(f"Warning: Could not load/apply qdarkstyle: {e}")

    # Create and show the main window
    window = MarkdownConverterApp()
    window.show()

    # Start the Qt event loop
    sys.exit(app.exec())
