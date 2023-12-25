import sys
from PyQt6.QtWidgets import (QApplication, QMainWindow, QPushButton, QLabel, QTextEdit, QFileDialog, QComboBox, QVBoxLayout, QHBoxLayout, QWidget, QMessageBox)
from PyQt6.QtCore import QThread, pyqtSignal
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
        self.setWindowTitle("OpenTranscribe")
        self.setGeometry(100, 100, 600, 600)

        # Model selection
        self.model_var = QComboBox()
        self.model_var.addItems(["tiny", "base", "small", "medium", "large"])
        self.model_var.setCurrentText("base")

        # File path label
        self.file_label = QLabel("No file selected")
        self.file_label.setStyleSheet("color: blue")

        # Buttons
        self.open_file_btn = QPushButton("Open File")
        self.open_file_btn.clicked.connect(self.select_file)
        self.transcribe_btn = QPushButton("Transcribe")
        self.transcribe_btn.clicked.connect(self.transcribe_audio)
        
        # Transcription text area
        self.transcription_text = QTextEdit()

        # Layout
        layout = QVBoxLayout()
        layout.addWidget(self.model_var)
        layout.addWidget(self.open_file_btn)
        layout.addWidget(self.file_label)
        layout.addWidget(self.transcribe_btn)
        layout.addWidget(self.transcription_text)

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
        self.transcription_thread = TranscriptionThread(file_path)
        self.transcription_thread.finished.connect(self.on_transcription_complete)
        self.transcription_thread.start()

    def on_transcription_complete(self, text):
        self.transcription_text.setText(text)

    def load_model(self, model_name):
        global current_model, current_model_name
        try:
            device = torch.device("mps") if torch.backends.mps.is_available() else torch.device("cpu")
            if model_name != current_model_name or current_model is None:
                current_model_name = model_name
                current_model = whisper.load_model(model_name).to(device)
        except NotImplementedError as e:
            QMessageBox.information(self, "Error", f"MPS not supported for this operation, falling back to CPU. Error {e}")
            device = torch.device("cpu")
            current_model = whisper.load_model(model_name).to(device)

# Main function
def main():
    app = QApplication(sys.argv)
    mainWindow = OpenTranscribe()
    mainWindow.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()