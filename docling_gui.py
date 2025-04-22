import sys
import os
import platform
import time # For simulating load if needed

# --- FIX for Hugging Face Hub Symlink Error on Windows ---
if platform.system() == 'Windows':
    print("Applying Windows Hugging Face Hub symlink workaround...")
    os.environ['HF_HUB_DISABLE_SYMLINKS'] = '1'
    os.environ['HF_HUB_DISABLE_SYMLINKS_WARNING'] = '1'
# --- End Fix ---

# --- GUI Imports (Place after potential env var changes) ---
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QFileDialog, QTextEdit, QLabel, QMessageBox, QSizePolicy # Added QSizePolicy
)
from PySide6.QtCore import Qt, Slot, QUrl, QTimer, QSize
from PySide6.QtGui import QDragEnterEvent, QDropEvent, QClipboard, QFont, QMovie # Added QMovie for potential GIF
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
# We will import inside the class initialization process


class MarkdownConverterApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setGeometry(100, 100, 800, 650)
        self.setWindowTitle("DocuMark Transformer") # Initial simple title

        self.original_status_text = "Initializing..." # Default status during load
        self.ready_status_text = "Ready. Select a file or drag it here."
        self.last_processed_file = None
        self.DoclingLoaderClass = None # Placeholder for the imported class

        # --- Central Widget & Layout ---
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        # Use a single main layout
        self.main_layout = QVBoxLayout(self.central_widget)
        self.main_layout.setContentsMargins(20, 20, 20, 15)
        self.main_layout.setSpacing(15)
        # Align content to center during loading, can be reset later if needed
        self.main_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        # --- Loading State Widget ---
        self.loading_label = QLabel("🚀 Initializing DocuMark Transformer...\n\nPlease wait, this might take a moment on first run.")
        loading_font = QFont()
        loading_font.setPointSize(14)
        self.loading_label.setFont(loading_font)
        self.loading_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.loading_label.setStyleSheet("color: #cccccc; padding: 50px;")
        # Ensure it can expand if needed
        self.loading_label.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.main_layout.addWidget(self.loading_label) # Add loading label first

        # --- Main UI Widgets (Created but initially hidden) ---
        self.title_label = QLabel("✨ DocuMark Transformer ✨")
        title_font = QFont()
        title_font.setPointSize(16)
        title_font.setBold(True)
        self.title_label.setFont(title_font)
        self.title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.title_label.setStyleSheet("padding-bottom: 10px; color: #5dade2;")

        self.open_button = QPushButton("📂 Open Document or HTML File")
        self.drop_label = QLabel("📄 ... or drag and drop a file here.")
        self.drop_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.markdown_output = QTextEdit()
        self.markdown_output.setReadOnly(True)
        self.markdown_output.setPlaceholderText("Converted Markdown will appear here...")

        self.button_layout_widget = QWidget() # Container for buttons
        self.button_layout = QHBoxLayout(self.button_layout_widget)
        self.copy_button = QPushButton("📋 Copy Markdown")
        self.save_button = QPushButton("💾 Save Markdown")
        self.copy_button.setEnabled(False)
        self.save_button.setEnabled(False)
        self.button_layout.addStretch(1)
        self.button_layout.addWidget(self.copy_button)
        self.button_layout.addWidget(self.save_button)
        self.button_layout.addStretch(1)
        # Remove margins from the button layout itself if needed
        self.button_layout.setContentsMargins(0, 0, 0, 0)

        self.status_label = QLabel(self.original_status_text)
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        # --- Apply Styles to hidden widgets ---
        button_min_height = 40
        button_padding = "8px 15px"
        border_radius = "5px"
        common_button_style = f"""
            QPushButton {{
                min-height: {button_min_height}px; padding: {button_padding}; border-radius: {border_radius}; font-size: 11pt;
            }}
            QPushButton:disabled {{ background-color: #4a4a4a; color: #888; }}
        """
        self.open_button.setStyleSheet(common_button_style)
        self.copy_button.setStyleSheet(common_button_style)
        self.save_button.setStyleSheet(common_button_style)
        icon_size = QSize(18, 18)
        self.open_button.setIconSize(icon_size)
        self.copy_button.setIconSize(icon_size)
        self.save_button.setIconSize(icon_size)

        self.base_drop_style = f"""
            QLabel {{ border: 2px dashed #666; padding: 40px 20px; border-radius: {border_radius}; background-color: #2a2a2a; font-size: 11pt; color: #aaa; }}
        """
        self.hover_drop_style = f"""
            QLabel {{ border: 2px dashed #bbb; padding: 40px 20px; border-radius: {border_radius}; background-color: #3a3a3a; font-size: 11pt; color: #eee; }}
        """
        self.drop_label.setStyleSheet(self.base_drop_style)

        self.markdown_output.setStyleSheet(f"""
            QTextEdit {{ border: 1px solid #555; border-radius: {border_radius}; padding: 10px; background-color: #282828; font-size: 10pt; }}
        """)
        self.status_label.setStyleSheet("QLabel { color: #999; padding-top: 5px; font-size: 9pt; }")

        # --- Add Widgets to Layout (but keep them hidden initially) ---
        # Loading label is already added and visible
        self.main_layout.addWidget(self.title_label)
        self.main_layout.addWidget(self.open_button)
        self.main_layout.addWidget(self.drop_label)
        self.main_layout.addWidget(self.markdown_output, 1) # Stretch factor
        self.main_layout.addWidget(self.button_layout_widget) # Add button container
        self.main_layout.addWidget(self.status_label) # Status label at the bottom

        # --- Hide Main UI Elements Initially ---
        self.title_label.setVisible(False)
        self.open_button.setVisible(False)
        self.drop_label.setVisible(False)
        self.markdown_output.setVisible(False)
        self.button_layout_widget.setVisible(False)
        # Keep status_label visible but it shows "Initializing..."

        # --- Enable Drag and Drop ---
        self.setAcceptDrops(True)

        # --- Connections (Connect signals even if widgets are hidden) ---
        self.open_button.clicked.connect(self.open_file_dialog)
        self.copy_button.clicked.connect(self.copy_markdown_to_clipboard)
        self.save_button.clicked.connect(self.save_markdown_to_file)
        self.markdown_output.textChanged.connect(self.update_action_buttons_state)

        # --- Start Initialization After GUI is Shown ---
        # Use a small delay (e.g., 100ms) to ensure the loading UI renders reliably
        QTimer.singleShot(100, self.initialize_backend)

    @Slot()
    def initialize_backend(self):
        """Performs the potentially slow import and setup."""
        print("Attempting to import DoclingLoader...")
        try:
            # Simulate potential longer load time (optional, for testing UI)
            # time.sleep(2)

            # --- Actual Import ---
            from langchain_docling import DoclingLoader as DL_Class
            self.DoclingLoaderClass = DL_Class # Store the imported class

            # Optional: Trigger a dummy load to ensure models are cached if needed
            # This might be useful if the very first conversion is slow otherwise.
            # Consider the implications (requires internet, disk space).
            # try:
            #     print("Performing dummy load to check cache...")
            #     # Create a dummy file path or use a known small one for testing
            #     # loader = self.DoclingLoaderClass(file_path="dummy.txt", export_type="markdown") # Needs valid path/file
            #     # loader.load() # This would trigger downloads if not cached
            #     print("Cache check complete (or skipped).")
            # except Exception as cache_err:
            #     print(f"Warning: Error during dummy load/cache check: {cache_err}")
            #     # Don't fail initialization for this, just log it.

            print("DoclingLoader imported successfully.")
            # Use QTimer to ensure UI update happens in the main event loop
            QTimer.singleShot(0, self.show_main_ui)

        except ImportError as e:
            print(f"Import Error: {e}")
            msg = f"Failed to import required library.\n\nError: {e}\n\nPlease ensure 'langchain-docling' and its dependencies are installed correctly.\n(pip install langchain-docling)"
            # Use QTimer for the error dialog as well
            QTimer.singleShot(0, lambda: self.show_initialization_error(msg))
        except Exception as e:
            # Catch other potential errors during import/setup
            import traceback
            print(f"Unexpected Initialization Error: {e}")
            traceback.print_exc()
            msg = f"An unexpected error occurred during initialization.\n\nError: {type(e).__name__}: {e}\n\nSee console output for more details."
            QTimer.singleShot(0, lambda: self.show_initialization_error(msg))

    @Slot()
    def show_main_ui(self):
        """Hides loading indicator and shows the main application UI."""
        print("Initialization complete. Showing main UI.")
        self.loading_label.setVisible(False)

        # Remove loading label from layout to allow proper resizing/alignment
        self.main_layout.removeWidget(self.loading_label)
        self.loading_label.deleteLater() # Clean up the widget

        # Reset alignment if it was changed for loading
        self.main_layout.setAlignment(Qt.AlignmentFlag.AlignTop) # Or remove alignment override

        # Show the main widgets
        self.title_label.setVisible(True)
        self.open_button.setVisible(True)
        self.drop_label.setVisible(True)
        self.markdown_output.setVisible(True)
        self.button_layout_widget.setVisible(True)

        self.setWindowTitle("DocuMark Transformer - Convert Documents to Markdown") # Set full title
        self.original_status_text = self.ready_status_text # Set the baseline status text
        self.set_status(self.original_status_text) # Update status label
        self.setEnabled(True) # Ensure window is interactive

    @Slot()
    def show_initialization_error(self, message):
        """Displays a critical error message if initialization fails."""
        print(f"Displaying initialization error: {message}")
        # Hide loading label even on error
        self.loading_label.setVisible(False)
        # Optionally display error in status bar too
        self.set_status("❌ Initialization Failed!")
        self.status_label.setStyleSheet("color: #e74c3c;") # Error color

        QMessageBox.critical(self, "Initialization Error", message)
        # Keep the window open but disable interactions as it's unusable
        self.setEnabled(False)


    # --- Event Handlers ---
    def dragEnterEvent(self, event: QDragEnterEvent):
        # Only accept drags if the main UI is visible and initialized
        if self.DoclingLoaderClass and self.open_button.isVisible():
            if event.mimeData().hasUrls():
                event.acceptProposedAction()
                self.drop_label.setStyleSheet(self.hover_drop_style)
            else:
                event.ignore()
        else:
            event.ignore() # Ignore drags during loading or if failed

    def dragLeaveEvent(self, event):
        # Only reset style if main UI is visible
        if self.DoclingLoaderClass and self.drop_label.isVisible():
            self.drop_label.setStyleSheet(self.base_drop_style)
        event.accept()

    def dropEvent(self, event: QDropEvent):
         # Only process drops if the main UI is visible and initialized
        if not (self.DoclingLoaderClass and self.open_button.isVisible()):
            event.ignore()
            return

        self.drop_label.setStyleSheet(self.base_drop_style) # Reset style first
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            urls = event.mimeData().urls()
            if urls:
                url = urls[0]
                file_path = ""
                # (Keep macOS path resolution logic as before)
                if url.isLocalFile():
                    file_path = url.toLocalFile()
                elif MACOS_DRAG_DROP_WORKAROUND and url.scheme() == 'file':
                    try:
                        ns_url = NSURL.URLWithString_(url.toString())
                        if ns_url and ns_url.isFileURL():
                            file_path = str(ns_url.path())
                        else:
                            self.show_error(f"Cannot resolve non-file URL on macOS: {url.toString()}")
                            return
                    except Exception as e:
                        self.show_error(f"Error resolving macOS path: {e}\nURL: {url.toString()}")
                        return
                else:
                    self.show_error(f"Cannot handle non-local file URL: {url.toString()}")
                    return

                if file_path:
                    self.process_file(file_path)
        else:
            event.ignore()


    @Slot()
    def open_file_dialog(self):
        # This should only be clickable if the main UI is visible
        if not self.DoclingLoaderClass: return # Extra safety check

        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Document or HTML File", "",
            "All Supported Files (*.pdf *.docx *.pptx *.html *.htm);;PDF Files (*.pdf);;Word Documents (*.docx);;PowerPoint Files (*.pptx);;HTML Files (*.html *.htm);;All Files (*)"
        )
        if file_path:
            self.process_file(file_path)

    def process_file(self, file_path: str):
        # Ensure initialization is complete before processing
        if not self.DoclingLoaderClass:
            self.show_error("Application is not fully initialized.")
            return

        if not os.path.exists(file_path):
            self.show_error(f"File not found: {file_path}")
            self.reset_status("File access error.")
            return
        if not os.access(file_path, os.R_OK):
             self.show_error(f"Permission denied: Cannot read file\n{file_path}")
             self.reset_status("File permission error.")
             return

        self.last_processed_file = file_path
        base_name = os.path.basename(file_path)
        self.set_status(f"⏳ Converting '{base_name}'...", is_processing=True)
        self.markdown_output.clear()
        self.update_action_buttons_state()
        QApplication.processEvents()

        try:
            # Use the stored class
            loader = self.DoclingLoaderClass(
                file_path=file_path,
                export_type="markdown",
            )
            docs = loader.load()

            # (Rest of the processing logic remains the same)
            if docs:
                page_contents = [doc.page_content for doc in docs if doc.page_content]
                if page_contents:
                    full_markdown = "\n\n".join(page_contents)
                    self.markdown_output.setText(full_markdown)
                    self.set_status(f"✅ Successfully converted '{base_name}'!", is_success=True)
                else:
                    self.show_error(f"Conversion resulted in empty content for '{base_name}'.")
                    self.reset_status("Conversion failed: Empty result.")
            else:
                self.show_error(f"Docling returned no processable documents for '{base_name}'.")
                self.reset_status("Conversion failed: No documents.")

        except OSError as e:
             # (Error handling remains the same)
             error_message = f"OS Error during conversion: {type(e).__name__}: {e}"
             print(error_message)
             if platform.system() == 'Windows' and isinstance(e, OSError) and hasattr(e, 'winerror') and e.winerror == 1314:
                 user_msg = f"Failed to convert '{base_name}'.\n\nPrivilege Error (WinError 1314).\n\nTroubleshooting:\n- Enable Developer Mode.\n- Run as Administrator.\n- Check cache permissions: C:\\Users\\{os.getlogin()}\\.cache\\huggingface"
             else:
                 user_msg = f"Failed to convert '{base_name}' due to an OS error.\n\nDetails: {e}\n\nCheck permissions and if file is open."
             self.show_error(user_msg)
             self.reset_status("❌ Conversion failed (OS Error).")
             self.markdown_output.clear()
        except ImportError as e:
             # (Error handling remains the same)
             error_message = f"Import Error during conversion: {type(e).__name__}: {e}"
             print(error_message)
             user_msg = f"Failed to convert '{base_name}'. Missing dependency?\n\nDetails: {e}"
             self.show_error(user_msg)
             self.reset_status("❌ Conversion failed (Missing Dependency).")
             self.markdown_output.clear()
        except Exception as e:
             # (Error handling remains the same)
             error_message = f"Unexpected error during conversion: {type(e).__name__}: {e}"
             import traceback
             print(error_message)
             traceback.print_exc()
             self.show_error(f"Failed to convert '{base_name}' (Unexpected Error).\n\nDetails: {type(e).__name__}\n\nSee console.")
             self.reset_status("❌ Conversion failed (Unexpected Error).")
             self.markdown_output.clear()
        finally:
            self.update_action_buttons_state()


    @Slot()
    def copy_markdown_to_clipboard(self):
        # (Logic remains the same)
        markdown_text = self.markdown_output.toPlainText()
        if markdown_text:
            try:
                clipboard = QApplication.clipboard()
                clipboard.setText(markdown_text)
                self.set_status("📋 Markdown copied to clipboard!", is_success=True, temporary=True)
            except Exception as e:
                self.show_error(f"Could not copy to clipboard: {e}")
                self.set_status("❌ Clipboard copy failed.", temporary=True)
        else:
            self.set_status("🤷‍♀️ Nothing to copy.", temporary=True)

    @Slot()
    def save_markdown_to_file(self):
        # (Logic remains the same)
        markdown_text = self.markdown_output.toPlainText()
        if not markdown_text:
            self.set_status("🤷‍♀️ Nothing to save.", temporary=True)
            return

        default_filename = "output.txt"
        if self.last_processed_file:
            base = os.path.basename(self.last_processed_file)
            name_without_ext = os.path.splitext(base)[0]
            default_filename = f"{name_without_ext}.txt"

        file_path, _ = QFileDialog.getSaveFileName(
            self, "Save Markdown As", default_filename, "Text Files (*.txt);;All Files (*)"
        )

        if file_path:
            try:
                if not file_path.lower().endswith(".txt"):
                    file_path += ".txt"
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(markdown_text)
                saved_filename = os.path.basename(file_path)
                self.set_status(f"💾 Saved to '{saved_filename}'", is_success=True, temporary=True)
            except OSError as e:
                self.show_error(f"Could not save file: {e}\n\nCheck directory permissions.")
                self.set_status("❌ File save failed (OS Error).", temporary=True)
            except Exception as e:
                self.show_error(f"Could not save file: {type(e).__name__}: {e}")
                self.set_status("❌ File save failed.", temporary=True)

    @Slot()
    def update_action_buttons_state(self):
        # (Logic remains the same)
        has_text = bool(self.markdown_output.toPlainText().strip())
        self.copy_button.setEnabled(has_text)
        self.save_button.setEnabled(has_text)

    def set_status(self, message: str, is_success=False, is_processing=False, temporary=False):
        # Use self.ready_status_text as the default non-temporary baseline
        if not is_processing and not temporary:
             self.original_status_text = message # Update the baseline text only if it's a non-transient state like "Ready" or "Converted"

        self.status_label.setText(message)
        # Reset style before applying specific ones
        self.status_label.setStyleSheet("color: #999; padding-top: 5px; font-size: 9pt;")
        # Optional styling based on state:
        # if is_success: self.status_label.setStyleSheet("color: #2ecc71; padding-top: 5px; font-size: 9pt;")
        # if is_processing: self.status_label.setStyleSheet("color: #f39c12; padding-top: 5px; font-size: 9pt;")

        if temporary:
            # Reset to the *current* baseline status after delay
            current_baseline = self.original_status_text
            QTimer.singleShot(3000, lambda: self.status_label.setText(current_baseline) if self.status_label.text() == message else None)


    def reset_status(self, base_message=None):
        # Resets status to the standard "Ready" message
        if base_message is None:
             base_message = self.ready_status_text # Use the stored ready message
        self.original_status_text = base_message
        QTimer.singleShot(0, lambda: self.status_label.setText(self.original_status_text))

    def show_error(self, message: str):
        # Only show error popup if the main UI is active
        if self.isEnabled():
            QTimer.singleShot(0, lambda: QMessageBox.warning(self, "Error", message))
        else:
             print(f"Suppressed Error Popup (Window Disabled): {message}")


if __name__ == "__main__":
    app = QApplication(sys.argv)

    # Apply the dark stylesheet globally
    try:
        stylesheet = qdarkstyle.load_stylesheet(qt_api='pyside6')
        app.setStyleSheet(stylesheet)
    except Exception as e:
        print(f"Warning: Could not load/apply qdarkstyle: {e}")

    # Create and show the window - initialization happens internally now
    window = MarkdownConverterApp()
    window.show()

    sys.exit(app.exec())