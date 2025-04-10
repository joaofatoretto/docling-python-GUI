import sys
import os
import platform

# --- FIX for Hugging Face Hub Symlink Error on Windows ---
# Set environment variables BEFORE importing libraries that use Hugging Face Hub
# This attempts to prevent OSError [WinError 1314] by disabling symlink creation,
# which often requires admin privileges or Developer Mode on Windows.
# It falls back to duplicating files in the cache, potentially using more disk space.
if platform.system() == 'Windows':
    print("Applying Windows Hugging Face Hub symlink workaround...")
    os.environ['HF_HUB_DISABLE_SYMLINKS'] = '1'
    # Also disable the related warning message, as we are choosing this mode intentionally.
    os.environ['HF_HUB_DISABLE_SYMLINKS_WARNING'] = '1'
# --- End Fix ---

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QFileDialog, QTextEdit, QLabel, QMessageBox
)
from PySide6.QtCore import Qt, Slot, QUrl, QTimer, QSize
from PySide6.QtGui import QDragEnterEvent, QDropEvent, QClipboard, QFont
import qdarkstyle # For dark theme


# Workaround for macOS drag-and-drop paths if needed
if platform.system() == 'Darwin':
    try:
        from Foundation import NSURL
        MACOS_DRAG_DROP_WORKAROUND = True
    except ImportError:
        MACOS_DRAG_DROP_WORKAROUND = False
else:
    MACOS_DRAG_DROP_WORKAROUND = False


class MarkdownConverterApp(QMainWindow):
    def __init__(self):
        super().__init__()
        # Increased default size for better spacing
        self.setGeometry(100, 100, 800, 650)

        self.original_status_text = "Ready. Select a file or drag it here." # Store initial status text

        # --- Central Widget & Layout ---
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.main_layout = QVBoxLayout(self.central_widget) # Main vertical layout

        # --- Styling Adjustments ---
        # Add overall padding and spacing to the main layout
        self.main_layout.setContentsMargins(20, 20, 20, 15) # L, T, R, B padding
        self.main_layout.setSpacing(15) # Spacing between widgets

        # --- Creative Title ---
        self.title_label = QLabel("✨ DocuMark Transformer ✨")
        title_font = QFont()
        title_font.setPointSize(16)
        title_font.setBold(True)
        self.title_label.setFont(title_font)
        self.title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.title_label.setStyleSheet("padding-bottom: 10px; color: #5dade2;") # Add some space below title and color

        # --- Widgets ---
        self.open_button = QPushButton("📂 Open Document or HTML File")
        self.drop_label = QLabel("📄 ... or drag and drop a file here.")
        self.drop_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.markdown_output = QTextEdit()
        self.markdown_output.setReadOnly(True)
        self.markdown_output.setPlaceholderText("Converted Markdown will appear here...")

        # --- Button Layout (Horizontal for Copy Button) ---
        self.button_layout = QHBoxLayout()
        self.copy_button = QPushButton("📋 Copy Markdown")
        self.copy_button.setEnabled(False) # Initially disabled

        # Center the copy button in its own layout area
        self.button_layout.addStretch(1)
        self.button_layout.addWidget(self.copy_button)
        self.button_layout.addStretch(1)

        # --- Status Label ---
        self.status_label = QLabel(self.original_status_text)
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        # --- Apply Styles ---
        button_min_height = 40 # Make buttons taller
        button_padding = "8px 15px" # Horizontal and vertical padding inside buttons
        border_radius = "5px" # Slightly rounded corners

        common_button_style = f"""
            QPushButton {{
                min-height: {button_min_height}px;
                padding: {button_padding};
                border-radius: {border_radius};
                font-size: 11pt; /* Slightly larger font */
            }}
            QPushButton:disabled {{
                background-color: #4a4a4a; /* Darker background when disabled */
                color: #888;
            }}
        """
        self.open_button.setStyleSheet(common_button_style)
        self.copy_button.setStyleSheet(common_button_style)
        # Add an icon size hint for buttons if needed (optional)
        self.open_button.setIconSize(QSize(18, 18))
        self.copy_button.setIconSize(QSize(18, 18))


        # Style for the drop label
        self.base_drop_style = f"""
            QLabel {{
                border: 2px dashed #666; /* Dashed border */
                padding: 40px 20px;      /* Generous vertical padding */
                border-radius: {border_radius};
                background-color: #2a2a2a; /* Slightly different background */
                font-size: 11pt;
                color: #aaa;
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

        # Style for the text output area
        self.markdown_output.setStyleSheet(f"""
            QTextEdit {{
                border: 1px solid #555;
                border-radius: {border_radius};
                padding: 10px;
                background-color: #282828; /* Slightly different bg */
                font-size: 10pt;
            }}
        """)

        # Style for the status label
        self.status_label.setStyleSheet("""
            QLabel {
                color: #999; /* Slightly muted color */
                padding-top: 5px;
                font-size: 9pt;
            }
        """)


        # --- Layout Setup ---
        self.main_layout.addWidget(self.title_label) # Add title first
        self.main_layout.addWidget(self.open_button)
        self.main_layout.addWidget(self.drop_label)
        self.main_layout.addWidget(self.markdown_output, 1) # Give text area stretch factor
        self.main_layout.addLayout(self.button_layout) # Add the horizontal layout for the button
        self.main_layout.addWidget(self.status_label) # Status at the bottom

        # --- Window Title (Still useful) ---
        self.setWindowTitle("DocuMark Transformer - Convert Documents to Markdown")

        # --- Enable Drag and Drop ---
        self.setAcceptDrops(True) # Accept drops on the main window

        # --- Connections ---
        self.open_button.clicked.connect(self.open_file_dialog)
        self.copy_button.clicked.connect(self.copy_markdown_to_clipboard)
        self.markdown_output.textChanged.connect(self.update_copy_button_state)

    # --- Event Handlers (Modified for Styling) ---

    def dragEnterEvent(self, event: QDragEnterEvent):
        # Accept drops only if they contain URLs (files)
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            # Apply hover style to the drop label
            self.drop_label.setStyleSheet(self.hover_drop_style)
        else:
            event.ignore()

    def dragLeaveEvent(self, event):
        # Restore base style when drag leaves the window
        self.drop_label.setStyleSheet(self.base_drop_style)
        event.accept() # Ensure event is handled

    def dropEvent(self, event: QDropEvent):
        # Restore base style after drop
        self.drop_label.setStyleSheet(self.base_drop_style)
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            urls = event.mimeData().urls()
            if urls:
                url = urls[0]
                file_path = ""
                if url.isLocalFile():
                    file_path = url.toLocalFile()
                elif MACOS_DRAG_DROP_WORKAROUND and url.scheme() == 'file':
                     try:
                         # Use the NSURL workaround for macOS Ventura+ sandboxing
                         ns_url = NSURL.URLWithString_(url.toString())
                         # Ensure it's a file path URL before getting the path
                         if ns_url and ns_url.isFileURL():
                             file_path = str(ns_url.path()) # Correct way to get path
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
                    # Process if dropped anywhere in the window
                    self.process_file(file_path)
        else:
            event.ignore()


    @Slot()
    def open_file_dialog(self):
        # Consider adding a check here to ensure models might be available
        # although the error happens during caching/first load anyway.
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Document or HTML File", "",
            "All Supported Files (*.pdf *.docx *.pptx *.html *.htm);;PDF Files (*.pdf);;Word Documents (*.docx);;PowerPoint Files (*.pptx);;HTML Files (*.html *.htm);;All Files (*)"
        )
        if file_path:
            self.process_file(file_path)

    def process_file(self, file_path: str):
        # Basic file existence and readability check
        if not os.path.exists(file_path):
            self.show_error(f"File not found: {file_path}")
            self.reset_status("File access error.")
            return
        if not os.access(file_path, os.R_OK):
             # Use os.access for read permission check before opening
             self.show_error(f"Permission denied: Cannot read file\n{file_path}")
             self.reset_status("File permission error.")
             return

        base_name = os.path.basename(file_path)
        self.set_status(f"⏳ Converting '{base_name}'...", is_processing=True)
        self.markdown_output.clear()
        self.copy_button.setEnabled(False) # Disable copy during processing
        QApplication.processEvents() # Update UI immediately

        try:
            # This is where DoclingLoader might trigger model downloads/caching
            # The HF_HUB_DISABLE_SYMLINKS=1 environment variable should prevent
            # the OSError [WinError 1314] during this process.
            loader = DoclingLoader(
                file_path=file_path,
                export_type="markdown",
                # md_export_kwargs={"include_images": False} # Example
            )
            # The .load() performs the actual conversion using the (now cached) models
            docs = loader.load()

            if docs:
                # Filter out potential None values in page_content just in case
                page_contents = [doc.page_content for doc in docs if doc.page_content]
                if page_contents:
                    full_markdown = "\n\n".join(page_contents)
                    self.markdown_output.setText(full_markdown)
                    self.set_status(f"✅ Successfully converted '{base_name}'!", is_success=True)
                else:
                    # Handle case where docs exist but have no content
                    self.show_error(f"Conversion resulted in empty content for '{base_name}'. The document might be empty or unsupported.")
                    self.reset_status("Conversion failed: Empty result.")
            else:
                # Handle case where loader returns no documents
                self.show_error(f"Docling returned no processable documents for '{base_name}'. It might be corrupted or an unsupported format/structure.")
                self.reset_status("Conversion failed: No documents.")

        # --- More Specific Error Handling ---
        except OSError as e:
             # Catch OS-level errors specifically, including potential permission issues
             # not caught by the initial check or file locking problems.
             error_message = f"OS Error during conversion: {type(e).__name__}: {e}"
             print(error_message) # Log detailed error
             # Check if it's the specific privilege error, although the env var should prevent it
             if isinstance(e, OSError) and e.winerror == 1314:
                 user_msg = f"Failed to convert '{base_name}'.\n\nA required file operation failed due to insufficient privileges (WinError 1314).\n\nTroubleshooting:\n- Ensure Developer Mode is enabled on Windows.\n- Try running the application as Administrator.\n- Check permissions for the cache folder:\n C:\\Users\\{os.getlogin()}\\.cache\\huggingface"
             else:
                 user_msg = f"Failed to convert '{base_name}' due to an OS error.\n\nDetails: {e}\n\nCheck file access permissions and ensure the file is not open elsewhere."
             self.show_error(user_msg)
             self.reset_status("❌ Conversion failed (OS Error).")
             self.markdown_output.clear()

        except ImportError as e:
            # Catch potential missing dependencies needed by Docling for specific file types
             error_message = f"Import Error during conversion: {type(e).__name__}: {e}"
             print(error_message)
             user_msg = f"Failed to convert '{base_name}'.\nA required dependency might be missing.\n\nDetails: {e}\n\nPlease check the Docling installation and requirements for handling this file type."
             self.show_error(user_msg)
             self.reset_status("❌ Conversion failed (Missing Dependency).")
             self.markdown_output.clear()

        except Exception as e:
             # Catch any other unexpected errors during the Docling process
             error_message = f"Unexpected error during conversion: {type(e).__name__}: {e}"
             import traceback
             print(error_message)
             traceback.print_exc() # Print full traceback for debugging
             self.show_error(f"Failed to convert '{base_name}' due to an unexpected error.\n\nDetails: {type(e).__name__}\n\nSee console/log for more details.")
             self.reset_status("❌ Conversion failed (Unexpected Error).")
             self.markdown_output.clear()
        # ------------------------------------

        finally:
            # This block always runs, ensuring the UI is re-enabled correctly
            self.update_copy_button_state() # Update based on whether output exists now


    @Slot()
    def copy_markdown_to_clipboard(self):
        """Copies the content of the markdown_output QTextEdit to the clipboard."""
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
    def update_copy_button_state(self):
        """Enables or disables the copy button based on whether there is text."""
        has_text = bool(self.markdown_output.toPlainText().strip())
        self.copy_button.setEnabled(has_text)

    def set_status(self, message: str, is_success=False, is_processing=False, temporary=False):
        """Updates the status label and potentially its style."""
        self.status_label.setText(message)
        # Storing original status only for non-temporary, non-processing states
        if not temporary and not is_processing:
            self.original_status_text = message

        if temporary:
            # Reset status label after a delay back to the stored original text
            QTimer.singleShot(2500, lambda: self.status_label.setText(self.original_status_text))


    def reset_status(self, base_message="Ready. Select a file or drag it here."):
        """Resets the status to a base message safely."""
        self.original_status_text = base_message
        # Use a timer to ensure this update happens safely in the Qt event loop
        QTimer.singleShot(0, lambda: self.status_label.setText(self.original_status_text))


    def show_error(self, message: str):
        """Shows a warning message box."""
        # Use a timer to ensure this dialog is shown safely from the Qt event loop
        QTimer.singleShot(0, lambda: QMessageBox.warning(self, "Error", message))


if __name__ == "__main__":
    # Set environment variables *first*
    if platform.system() == 'Windows':
        print("Applying Windows Hugging Face Hub symlink workaround (main block)...")
        os.environ['HF_HUB_DISABLE_SYMLINKS'] = '1'
        os.environ['HF_HUB_DISABLE_SYMLINKS_WARNING'] = '1'

    app = QApplication(sys.argv)

    # Apply the dark stylesheet globally
    try:
        stylesheet = qdarkstyle.load_stylesheet(qt_api='pyside6')
        app.setStyleSheet(stylesheet)
    except Exception as e:
        print(f"Warning: Could not load/apply qdarkstyle: {e}")
        # App will continue with default styling

    window = MarkdownConverterApp()
    window.show()
    
    # Import Docling components via LangChain integration
    # This import should happen AFTER the environment variables are set
    try:
        from langchain_docling import DoclingLoader
    except ImportError:
        print("Error: langchain-docling not found. Please install it: pip install langchain-docling")
        sys.exit(1)
    except Exception as e:
        print(f"Error importing langchain-docling: {e}") # Catch other potential import errors
        # Display error in a GUI way if possible, or exit.
        # Since QApplication might not be running yet, print and exit is safer here.
        sys.exit(1)
    
    sys.exit(app.exec())