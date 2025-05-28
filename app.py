import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.ttk import Progressbar
import os
from docx import Document

class XMLHighlighterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("XML to DOCX Batch Highlighter")
        self.ref_filepath = None
        self.target_files = []

        # Reference file selection
        ref_frame = tk.Frame(root, pady=5)
        ref_frame.pack(fill='x')
        tk.Button(ref_frame, text="Select Reference DOCX", command=self.load_reference).pack(side='left', padx=5)
        self.ref_label = tk.Label(ref_frame, text="No reference file selected", wraplength=400)
        self.ref_label.pack(side='left', padx=5)

        # Target files selection
        tgt_frame = tk.Frame(root, pady=5)
        tgt_frame.pack(fill='x')
        tk.Button(tgt_frame, text="Select XML Files", command=self.load_targets).pack(side='left', padx=5)
        self.tgt_label = tk.Label(tgt_frame, text="No XML files selected", wraplength=400)
        self.tgt_label.pack(side='left', padx=5)

        # Output folder selection
        out_frame = tk.Frame(root, pady=5)
        out_frame.pack(fill='x')
        tk.Button(out_frame, text="Select Output Folder", command=self.load_output_folder).pack(side='left', padx=5)
        self.out_label = tk.Label(out_frame, text="No output folder selected", wraplength=400)
        self.out_label.pack(side='left', padx=5)

        # Progress bar
        self.progress = Progressbar(root, orient='horizontal', length=500, mode='determinate')
        self.progress.pack(pady=10)

        # Start button
        tk.Button(root, text="Start Highlighting", command=self.start).pack(pady=10)

    def load_reference(self):
        path = filedialog.askopenfilename(
            title="Select Reference DOCX", filetypes=[("Word files", "*.docx")]
        )
        if path:
            self.ref_filepath = path
            self.ref_label.config(text=os.path.basename(path))

    def load_targets(self):
        files = filedialog.askopenfilenames(
            title="Select XML Files to Analyze", filetypes=[("XML files", "*.xml")]
        )
        if files:
            self.target_files = list(files)
            self.tgt_label.config(text=f"{len(files)} files selected")

    def load_output_folder(self):
        folder = filedialog.askdirectory(title="Select Output Folder")
        if folder:
            self.output_folder = folder
            self.out_label.config(text=folder)

    def extract_reference_strings(self):
        ref_doc = Document(self.ref_filepath)
        strings = set()
        for para in ref_doc.paragraphs:
            text = para.text.strip()
            if text.startswith(('•', '-', '*')) or 'list paragraph' in getattr(para.style, 'name', '').lower():
                cleaned = text.lstrip('•-–* ').strip()
                if cleaned:
                    strings.add(cleaned)
        return sorted(strings, key=len, reverse=True)

    def highlight_in_docx(self, doc, text, ref_strings):
        runs = []
        idx = 0
        while idx < len(text):
            match = None
            for ref in ref_strings:
                if text.startswith(ref, idx):
                    match = ref
                    break
            if match:
                runs.append({'text': match, 'highlight': True})
                idx += len(match)
            else:
                runs.append({'text': text[idx], 'highlight': False})
                idx += 1
        para = self.current_paragraph
        for run in runs:
            r = para.add_run(run['text'])
            if run['highlight']:
                r.font.highlight_color = 7

    def process_file(self, xml_path, ref_strings):
        with open(xml_path, 'r', encoding='utf-8') as f:
            lines = f.readlines()
        doc = Document()
        # Remove default empty content
        body = doc._body._element
        for child in list(body):
            body.remove(child)
        for line in lines:
            para = doc.add_paragraph()
            self.current_paragraph = para
            self.highlight_in_docx(doc, line.rstrip('\n'), ref_strings)
        out_name = os.path.splitext(os.path.basename(xml_path))[0] + '_highlighted.docx'
        doc.save(os.path.join(self.output_folder, out_name))

    def start(self):
        if not self.ref_filepath or not self.target_files or not hasattr(self, 'output_folder'):
            messagebox.showerror("Error", "Please select a reference DOCX, XML files, and output folder.")
            return
        ref_strings = self.extract_reference_strings()
        self.progress['maximum'] = len(self.target_files)
        for i, xml_file in enumerate(self.target_files, 1):
            self.process_file(xml_file, ref_strings)
            self.progress['value'] = i
            self.root.update_idletasks()
        messagebox.showinfo("Done", "Batch highlighting complete!")

if __name__ == '__main__':
    root = tk.Tk()
    app = XMLHighlighterApp(root)
    root.mainloop()
