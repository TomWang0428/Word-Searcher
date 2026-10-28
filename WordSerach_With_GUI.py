import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, Toplevel
from docx import Document
import openpyxl
from pptx import Presentation
import Levenshtein
import fitz  # PyMuPDF for PDF support

# ---------------- Matching Logic ------------------
def is_match(word, keyword, max_mismatches, use_substring_match, use_blurry):
    word_lower = word.lower()
    keyword_lower = keyword.lower()

    if use_substring_match and keyword_lower in word_lower:
        return True

    if use_blurry:
        return Levenshtein.distance(word_lower, keyword_lower) <= max_mismatches

    return word_lower == keyword_lower

# ---------------- Text Extraction ------------------
def extract_text_from_docx(path):
    try:
        doc = Document(path)
        return "\n".join(para.text for para in doc.paragraphs)
    except:
        return ""

def extract_text_from_xlsx(path):
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
        text = ""
        for sheet in wb.worksheets:
            for row in sheet.iter_rows():
                for cell in row:
                    if cell.value:
                        text += str(cell.value) + " "
        return text
    except:
        return ""

def extract_text_from_pptx(path):
    try:
        prs = Presentation(path)
        text = ""
        for slide in prs.slides:
            for shape in slide.shapes:
                if hasattr(shape, "text"):
                    text += shape.text + " "
        return text
    except:
        return ""

def extract_text_from_pdf(path):
    try:
        doc = fitz.open(path)
        text = ""
        for page in doc:
            text += page.get_text()
        return text
    except:
        return ""

def extract_text(path):
    if path.endswith(".docx"):
        return extract_text_from_docx(path)
    elif path.endswith(".xlsx"):
        return extract_text_from_xlsx(path)
    elif path.endswith(".pptx"):
        return extract_text_from_pptx(path)
    elif path.endswith(".pdf"):
        return extract_text_from_pdf(path)
    return ""

# ---------------- GUI App ------------------
class SearchApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Document Keyword Search")
        self.root.geometry("1200x700")
        self.keywords = []

        # Frames
        self.left_frame = tk.Frame(root, width=200)
        self.left_frame.pack(side="left", fill="y")

        self.middle_frame = tk.Frame(root, width=300)
        self.middle_frame.pack(side="left", fill="y")

        self.right_frame = tk.Frame(root, width=500)
        self.right_frame.pack(side="left", fill="both", expand=True)

        # Left Panel - Directory + Keywords
        self.dir_label = tk.Label(self.left_frame, text="Directory:", font=("TkDefaultFont", 9, "bold"))
        self.dir_label.pack(anchor="w", padx=5)

        self.dir_entry = tk.Entry(self.left_frame, width=30)
        self.dir_entry.pack(padx=5, pady=2)

        self.browse_button = tk.Button(self.left_frame, text="Browse", command=self.browse_directory)
        self.browse_button.pack(pady=2)

        # File type filters
        self.filetypes_frame = tk.Frame(self.left_frame)
        self.filetypes_frame.pack(pady=2)
        self.include_docx = tk.BooleanVar(value=True)
        self.include_xlsx = tk.BooleanVar(value=True)
        self.include_pptx = tk.BooleanVar(value=True)
        self.include_pdf = tk.BooleanVar(value=True)
        tk.Checkbutton(self.filetypes_frame, text="Word (.docx)", variable=self.include_docx).pack(anchor="w")
        tk.Checkbutton(self.filetypes_frame, text="Excel (.xlsx)", variable=self.include_xlsx).pack(anchor="w")
        tk.Checkbutton(self.filetypes_frame, text="PowerPoint (.pptx)", variable=self.include_pptx).pack(anchor="w")
        tk.Checkbutton(self.filetypes_frame, text="PDF (.pdf)", variable=self.include_pdf).pack(anchor="w")

        # Search options title
        tk.Label(self.left_frame, text="Search Options:", font=("TkDefaultFont", 9, "bold")).pack(anchor="w", padx=5, pady=5)

        # Blurry + Substring options grouped
        self.search_opts_frame = tk.Frame(self.left_frame)
        self.search_opts_frame.pack(pady=5, anchor="w")

        blurry_row = tk.Frame(self.search_opts_frame)
        blurry_row.pack(anchor="w", padx=5)

        self.use_blurry = tk.BooleanVar(value=True)
        tk.Checkbutton(blurry_row, text="Blurry search", variable=self.use_blurry).pack(side="left")

        self.max_mismatch_entry = tk.Entry(blurry_row, width=5)
        self.max_mismatch_entry.insert(0, "1")
        self.max_mismatch_entry.pack(side="left", padx=(5, 0))
        self.use_substring_match = tk.BooleanVar(value=True)
        tk.Checkbutton(self.search_opts_frame, text="Allow partial match", variable=self.use_substring_match).pack(anchor="w", padx=5)

        self.search_button = tk.Button(self.left_frame, text="Search", command=self.search_documents)
        self.search_button.pack(pady=10)

        self.keyword_label = tk.Label(self.left_frame, text="Keywords:", font=("TkDefaultFont", 9, "bold"))
        self.keyword_label.pack(anchor="w", padx=5)

        self.keyword_control_frame = tk.Frame(self.left_frame)
        self.keyword_control_frame.pack(padx=5, pady=2, anchor="center")
        self.add_button = tk.Button(self.keyword_control_frame, text="+", width=5, command=self.add_keyword_entry)
        self.add_button.pack(side="left", padx=2)
        self.remove_button = tk.Button(self.keyword_control_frame, text="-", width=5, command=self.remove_keyword_entry)
        self.remove_button.pack(side="left", padx=2)

        self.keyword_entries = []
        self.add_keyword_entry()

        # Middle Panel - Result List
        self.result_listbox = tk.Listbox(self.middle_frame, width=60)
        self.result_listbox.pack(fill="both", expand=True)
        self.result_listbox.bind('<<ListboxSelect>>', self.on_file_select)

        # Right Panel - Match Context
        self.context_text = tk.Text(self.right_frame, wrap="word")
        self.context_text.pack(fill="both", expand=True)

        self.results = {}

    def add_keyword_entry(self):
        entry = tk.Entry(self.left_frame)
        entry.pack(padx=5, pady=2)
        self.keyword_entries.append(entry)

    def remove_keyword_entry(self):
        if self.keyword_entries:
            entry = self.keyword_entries.pop()
            entry.destroy()

    def get_keywords(self):
        return [e.get().strip() for e in self.keyword_entries if e.get().strip() != ""]

    def browse_directory(self):
        directory = filedialog.askdirectory(title="Select Directory")
        if directory:
            self.dir_entry.delete(0, tk.END)
            self.dir_entry.insert(0, directory)

    def search_documents(self):
        directory = self.dir_entry.get().strip()
        if not directory or not os.path.isdir(directory):
            messagebox.showwarning("Invalid Directory", "Please select a valid directory.")
            return

        keywords = self.get_keywords()
        if not keywords:
            messagebox.showwarning("No Keywords", "Please enter at least one keyword.")
            return

        try:
            max_mismatches = int(self.max_mismatch_entry.get())
        except:
            max_mismatches = 1

        use_blurry = self.use_blurry.get()
        allow_partial = self.use_substring_match.get()

        filetypes = []
        if self.include_docx.get():
            filetypes.append(".docx")
        if self.include_xlsx.get():
            filetypes.append(".xlsx")
        if self.include_pptx.get():
            filetypes.append(".pptx")
        if self.include_pdf.get():
            filetypes.append(".pdf")
        if not filetypes:
            messagebox.showwarning("No File Types", "Please select at least one file type to search.")
            return

        self.result_listbox.delete(0, tk.END)
        self.context_text.delete(1.0, tk.END)
        self.results.clear()

        all_files = []
        for root, _, files in os.walk(directory):
            for file in files:
                if any(file.endswith(ext) for ext in filetypes):
                    all_files.append(os.path.join(root, file))

        # Progress popup
        progress_popup = Toplevel(self.root)
        progress_popup.title("Searching...")
        tk.Label(progress_popup, text="Searching files...").pack(padx=10, pady=5)
        progress = ttk.Progressbar(progress_popup, orient="horizontal", length=250, mode="determinate")
        progress.pack(padx=10, pady=10)
        progress["maximum"] = len(all_files)
        progress["value"] = 0
        self.root.update_idletasks()

        for idx, full_path in enumerate(all_files):
            text = extract_text(full_path)
            words = text.split()
            snippets = []

            for i, word in enumerate(words):
                for keyword in keywords:
                    if is_match(word, keyword, max_mismatches, allow_partial, use_blurry):
                        start = max(0, i - 3)
                        end = min(len(words), i + 4)
                        context = words[start:end]
                        snippets.append((" ".join(context), word))
                        break

            if snippets:
                self.results[full_path] = snippets
                self.result_listbox.insert(tk.END, full_path)

            progress["value"] = idx + 1
            self.root.update_idletasks()

        progress_popup.destroy()

    def on_file_select(self, event):
        selection = event.widget.curselection()
        if not selection:
            return

        index = selection[0]
        filepath = event.widget.get(index)
        snippets = self.results.get(filepath, [])

        self.context_text.delete(1.0, tk.END)

        for snippet_text, matched_word in snippets:
            words = snippet_text.split()
            highlighted = [f"[{w}]" if w.lower() == matched_word.lower() else w for w in words]
            self.context_text.insert(tk.END, f"... {' '.join(highlighted)} ...\n\n")

# ---------------- Run the App ------------------
if __name__ == "__main__":
    root = tk.Tk()
    app = SearchApp(root)
    root.mainloop()
