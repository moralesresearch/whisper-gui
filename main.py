import sys
from PyQt6.QtWidgets import (QApplication, QMainWindow, QPushButton, QLabel, QTextEdit, QFileDialog, QComboBox,
                             QVBoxLayout, QHBoxLayout, QWidget, QMessageBox, QFrame)
from PyQt6.QtCore import QThread, pyqtSignal, Qt
import whisper
import torch
from docx import Document

# Global variable for the Whisper model
current_model = None
current_model_name = "base"  # Default model size

class TranscriptionThread(QThread):
    finished = pyqtSignal(str)

    def __init__(self, file_path):
        super().__init__()
        self.file_path = file_path

    def run(self):
        global current_model
        try:
            result = current_model.transcribe(self.file_path)
            self.finished.emit(result["text"])
        except Exception as e:
            self.finished.emit(str(e))

class OpenTranscribe(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle("OpenTranscribe 3.0")
        self.setGeometry(100, 100, 600, 600)

        # Model selection
        self.model_var = QComboBox()
        self.model_var.addItems(["tiny", "base", "small", "medium", "large"])
        self.model_var.setCurrentText("small")

        # File path label
        self.file_label = QLabel("No file selected")
        self.file_label.setStyleSheet("color: blue")

        # Export Frame
        export_frame = QFrame()
        export_layout = QHBoxLayout()
        export_frame.setLayout(export_layout)

        # Export Buttons
        export_txt_btn = QPushButton("Export as TXT")
        export_txt_btn.clicked.connect(self.save_text)
        export_rtf_btn = QPushButton("Export as RTF")  # RTF might require additional handling
        export_rtf_btn.clicked.connect(self.save_rtf)
        export_docx_btn = QPushButton("Export as DOCX")
        export_docx_btn.clicked.connect(self.save_docx)

        export_layout.addWidget(export_txt_btn)
        export_layout.addWidget(export_rtf_btn)
        export_layout.addWidget(export_docx_btn)


        # Buttons
        self.open_file_btn = QPushButton("Open File")
        self.open_file_btn.clicked.connect(self.select_file)
        self.transcribe_btn = QPushButton("Transcribe")
        self.transcribe_btn.clicked.connect(self.transcribe_audio)
        
        # Transcription text area
        self.transcription_text = QTextEdit()

        # Status Label
        self.status_label = QLabel("")
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        # Layout
        layout = QVBoxLayout()
        layout.addWidget(self.model_var)
        layout.addWidget(self.open_file_btn)
        layout.addWidget(self.file_label)
        layout.addWidget(self.transcribe_btn)
        layout.addWidget(self.transcription_text)
        layout.addWidget(self.status_label)
        layout.addWidget(export_frame)

        # Main widget
        central_widget = QWidget()
        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)

    def select_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select File", "", "Audio Files (*.mp3 *.wav *.m4a)")
        if file_path:
            self.file_label.setText(file_path)

    def transcribe_audio(self):
        file_path = self.file_label.text()
        if file_path == "No file selected":
            QMessageBox.information(self, "Error", "No file selected")
            return
        selected_model = self.model_var.currentText()
        self.load_model(selected_model)
        # Update status to 'Transcribing...'
        self.status_label.setText("Transcribing...")
        self.status_label.setStyleSheet("color: black;")  # Default color
        self.transcription_thread = TranscriptionThread(file_path)
        self.transcription_thread.finished.connect(self.on_transcription_complete)
        self.transcription_thread.start()

    def on_transcription_complete(self, text):
        self.transcription_text.setText(text)
        # Update status to 'Done' and change color to green
        self.status_label.setText("Done")
        self.status_label.setStyleSheet("color: green;")

    def load_model(self, model_name):
        global current_model, current_model_name
        try:
            # First, try to set the device to MPS if available
            device = torch.device("mps") if torch.backends.mps.is_available() else torch.device("cpu")
            # Attempt to load the model on the MPS device
            if model_name != current_model_name or current_model is None:
                current_model_name = model_name
                current_model = whisper.load_model(model_name).to(device)
        except Exception as e:
            # If any exception occurs, fall back to CPU and display a message
            QMessageBox.information(self, "Info", "MPS is not yet supported by Whisper, falling back to CPU.")
            device = torch.device("cpu")
            current_model = whisper.load_model(model_name).to(device)

    def save_text(self):
        file_path, _ = QFileDialog.getSaveFileName(self, "Save File", "", "Text Files (*.txt)")
        if file_path:
            with open(file_path, "w") as file:
                file.write(self.transcription_text.toPlainText())

    def save_rtf(self):
        file_path, _ = QFileDialog.getSaveFileName(self, "Save File", "", "RTF Files (*.rtf)")
        if file_path:
            text = self.transcription_text.toPlainText()
            # RTF header and footer
            rtf_header = r"{\rtf1\ansi\ansicpg1252\deff0\nouicompat\deflang1033{\fonttbl{\f0\fnil\fcharset0 Calibri;}}"
            rtf_footer = r"}"
            # Converting plain text to RTF-compatible paragraph format
            rtf_body = text.replace('\n', '\\par\n')
            # Full RTF content
            rtf_content = rtf_header + rtf_body + rtf_footer
            with open(file_path, "w") as file:
                file.write(rtf_content)

    def save_docx(self):
        file_path, _ = QFileDialog.getSaveFileName(self, "Save File", "", "Word Documents (*.docx)")
        if file_path:
            doc = Document()
            doc.add_paragraph(self.transcription_text.toPlainText())
            doc.save(file_path)

# Main function
def main():
    app = QApplication(sys.argv)
    mainWindow = OpenTranscribe()
    mainWindow.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()