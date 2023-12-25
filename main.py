import tkinter as tk
import torch
from tkinter import filedialog, messagebox
import whisper
import threading
import gc  # Garbage collection
from docx import Document

# Global variable for the Whisper model
current_model = None
current_model_name = "base"  # Default model size

def load_model(model_name):
    global current_model, current_model_name
    try:
        device = torch.device("mps") if torch.backends.mps.is_available() else torch.device("cpu")
        if model_name != current_model_name or current_model is None:
            current_model_name = model_name
            current_model = whisper.load_model(model_name).to(device)
    except NotImplementedError as e:
        # Fallback if compat is not resolved
        print("MPS not supported for this operation, falling back to CPU. Error ", e)
        device = torch.device("cpu")
        current_model = whisper.load_model(model_name).to(device)
def select_file():
    file_path = filedialog.askopenfilename()
    file_label.config(text=file_path)
    return file_path

def transcribe_audio():
    file_path = file_label.cget("text")
    selected_model = model_var.get()
    if not file_path:
        messagebox.showinfo("Error", "No file selected")
        return
    load_model(selected_model)  # Load the selected model
    update_status("Transcribing...")
    transcribe_btn.config(state="disabled")
    threading.Thread(target=run_transcription, args=(file_path,), daemon=True).start()

def run_transcription(file_path):
    try:
        result = current_model.transcribe(file_path)
        transcription_text.delete("1.0", tk.END)
        transcription_text.insert(tk.END, result["text"])
    except Exception as e:
        messagebox.showerror("Error", str(e))
    finally:
        transcribe_btn.config(state="normal")
        update_status("")
        gc.collect()  # Explicitly call garbage collector

def update_status(message):
    status_label.config(text=message)
    root.update_idletasks()

def save_text():
    text = transcription_text.get("1.0", tk.END)
    save_file(text, "txt")

def save_rtf():
    text = transcription_text.get("1.0", tk.END)
    save_file(text, "rtf")

def save_docx():
    text = transcription_text.get("1.0", tk.END)
    save_file(text, "docx")

def save_file(text, file_type):
    file_path = filedialog.asksaveasfilename(defaultextension=f".{file_type}",
                                             filetypes=[(f"{file_type.upper()} files", f"*.{file_type}")])
    if file_path:
        if file_type == "txt" or file_type == "rtf":
            with open(file_path, "w") as file:
                file.write(text)
        elif file_type == "docx":
            doc = Document()
            doc.add_paragraph(text)
            doc.save(file_path)

# Create the main window
root = tk.Tk()
root.title("OpenTranscribe")
root.geometry("600x600")  # Set window size

# Model selection
model_var = tk.StringVar(root)
model_var.set("base")  # default value
models = ["tiny", "base", "small", "medium", "large"]
model_menu = tk.OptionMenu(root, model_var, *models)

# Create widgets
open_file_btn = tk.Button(root, text="Open File", command=select_file)
file_label = tk.Label(root, text="No file selected", fg="blue")
transcribe_btn = tk.Button(root, text="Transcribe", command=transcribe_audio)
transcription_text = tk.Text(root, wrap="word")

# Export buttons frame
export_frame = tk.Frame(root)
export_txt_btn = tk.Button(export_frame, text="Export as TXT", command=save_text)
export_rtf_btn = tk.Button(export_frame, text="Export as RTF", command=save_rtf)
export_docx_btn = tk.Button(export_frame, text="Export as DOCX", command=save_docx)

# Arrange widgets in the frame
export_txt_btn.grid(row=0, column=0, padx=5, pady=5)
export_rtf_btn.grid(row=0, column=1, padx=5, pady=5)
export_docx_btn.grid(row=0, column=2, padx=5, pady=5)

# Status label
status_label = tk.Label(root, text="", fg="red")
status_label.pack(anchor='se', padx=10, pady=5)

# Arrange other widgets
model_menu.pack(pady=5)
open_file_btn.pack(pady=5)
file_label.pack(pady=5)
transcribe_btn.pack(pady=5)
transcription_text.pack(expand=True, fill=tk.BOTH, padx=10, pady=10)
export_frame.pack(pady=5)

# Start the GUI event loop
root.mainloop()