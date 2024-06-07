import os
import threading
import time
import logging
from tkinter import Tk, filedialog, Button, Label, ttk, StringVar, messagebox, Checkbutton, BooleanVar
from PyPDF2 import PdfReader
from concurrent.futures import ProcessPoolExecutor, as_completed
import multiprocessing
import pandas as pd

# Setup logging
logging.basicConfig(filename='pdf_page_counter.log', level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')

# Constants
SEARCH_PROGRESS_UPDATE_INTERVAL = 10
PROCESSING_PROGRESS_UPDATE_INTERVAL = 10


class PDFPageCounterApp:
    def __init__(self, root):
        self.root = root
        self.folder_path = None
        self.stop_event = threading.Event()
        self.search_subfolders_var = BooleanVar(
            value=True)  # Variable for checkbox
        self.initialize_variables()
        self.setup_gui()

    def initialize_variables(self):
        self.folder_label_var = StringVar(value="No folder selected")
        self.search_progress_label_var = StringVar(
            value="Search Progress: 0/0 files")
        self.processing_progress_label_var = StringVar(
            value="Processing Progress: 0/0 files")
        self.results_var = StringVar(value="Results will be shown here.")
        self.execution_time_var = StringVar(
            value="Execution Time: Not yet run")
        self.status_label_var = StringVar(value="Status: Ready")
        self.search_progress_bar = None
        self.processing_progress_bar = None
        self.run_button = None

    def setup_gui(self):
        Label(self.root, text="Select a folder containing PDF files.",
              wraplength=400).pack(pady=10)
        Label(self.root, textvariable=self.folder_label_var).pack(pady=5)
        Button(self.root, text="Select Folder",
               command=self.select_folder).pack(pady=5)

        # Add checkbox for subfolder search
        Checkbutton(self.root, text="Search in subfolders",
                    variable=self.search_subfolders_var).pack(pady=5)

        self.run_button = Button(self.root, text="Start Counting Pages",
                                 command=self.toggle_script)
        self.run_button.pack(pady=20)
        self.search_progress_bar = self.create_progress_bar(
            self.search_progress_label_var)
        self.processing_progress_bar = self.create_progress_bar(
            self.processing_progress_label_var)
        Label(self.root, textvariable=self.results_var,
              justify="left").pack(pady=5)
        Label(self.root, textvariable=self.execution_time_var).pack(pady=5)
        Label(self.root, textvariable=self.status_label_var).pack(pady=5)

    def create_progress_bar(self, label_var):
        Label(self.root, textvariable=label_var).pack(pady=5)
        progress_bar = ttk.Progressbar(
            self.root, orient="horizontal", length=300, mode="determinate")
        progress_bar.pack(pady=5)
        return progress_bar

    def select_folder(self):
        self.folder_path = filedialog.askdirectory()
        self.folder_label_var.set(f"Selected Folder: {self.folder_path}")

    def toggle_script(self):
        if self.run_button['text'] == "Start Counting Pages":
            self.run_button['text'] = "Stop Script"
            self.run_button['command'] = self.stop_script
            self.stop_event.clear()
            start_time = time.time()
            threading.Thread(target=self.execute_task,
                             args=(start_time,)).start()
        else:
            self.stop_script()

    def stop_script(self):
        self.stop_event.set()
        self.run_button['text'] = "Start Counting Pages"
        self.run_button['command'] = self.toggle_script

    def execute_task(self, start_time):
        try:
            self.reset_progress_bars()

            self.validate_folder_path()
            self.update_status("Searching for PDF files...")
            pdf_files = self.search_pdf_files(
                self.folder_path, self.update_search_progress, self.search_subfolders_var.get())
            self.update_status("Counting pages in PDF files...")
            results, errors = self.count_pdf_pages(
                pdf_files, self.update_processing_progress)
            self.display_results(results)

            messagebox_content = None

            self.update_status("Saving results to Excel...")
            self.save_results_to_excel(self.folder_path, results, errors)
            messagebox_content = f"Excel report file saved to: {self.folder_path}"
            self.update_status("Ready")

            # Capture end time and update execution time
            execution_time = time.time() - start_time
            self.execution_time_var.set(
                f"Execution Time: {execution_time:.2f} seconds")
            self.root.update_idletasks()  # Ensure the GUI updates immediately

            # Show messagebox content after updating the execution time
            if messagebox_content:
                messagebox.showinfo("Info", messagebox_content)

        except Exception as e:
            logging.error("An error occurred", exc_info=True)
            self.execution_time_var.set("Execution Time: Error occurred")
            self.root.update_idletasks()  # Ensure the GUI updates immediately
            self.display_error(str(e))
            self.update_status("Error occurred")
        finally:
            self.run_button['text'] = "Start Counting Pages"
            self.run_button['command'] = self.toggle_script

    def validate_folder_path(self):
        if not self.folder_path:
            raise ValueError("Please select a folder.")

    def display_results(self, results):
        total_pages = sum(pages for _, pages, _,
                          _ in results if pages is not None)
        total_pdfs = len(results)
        self.results_var.set(
            f"Total Pages: {total_pages}\nTotal PDFs: {total_pdfs}")

    def update_search_progress(self, current, total):
        if self.stop_event.is_set():
            raise Exception("Script stopped by user.")
        self.search_progress_bar['maximum'] = total
        self.search_progress_bar['value'] = current
        self.search_progress_label_var.set(
            f"Search Progress: {current}/{total} files")
        self.root.update_idletasks()

    def update_processing_progress(self, current, total):
        if self.stop_event.is_set():
            raise Exception("Script stopped by user.")
        self.processing_progress_bar['maximum'] = total
        self.processing_progress_bar['value'] = current
        self.processing_progress_label_var.set(
            f"Processing Progress: {current}/{total} files")
        self.root.update_idletasks()

    def update_status(self, status):
        self.status_label_var.set(f"Status: {status}")
        self.root.update_idletasks()

    def display_error(self, error_message):
        messagebox.showerror("Error", error_message)

    @staticmethod
    def search_pdf_files(directory, progress_callback, search_subfolders):
        pdf_files = []
        if search_subfolders:
            total_files = sum(len(files) for _, _, files in os.walk(directory))
            processed_files = 0
            for root, _, files in os.walk(directory):
                for file in files:
                    processed_files += 1
                    if file.endswith('.pdf'):
                        pdf_files.append(os.path.join(root, file))
                    if processed_files % SEARCH_PROGRESS_UPDATE_INTERVAL == 0:
                        progress_callback(processed_files, total_files)
            progress_callback(total_files, total_files)
        else:
            total_files = len(os.listdir(directory))
            processed_files = 0
            for file in os.listdir(directory):
                processed_files += 1
                if file.endswith('.pdf'):
                    pdf_files.append(os.path.join(directory, file))
                if processed_files % SEARCH_PROGRESS_UPDATE_INTERVAL == 0:
                    progress_callback(processed_files, total_files)
            progress_callback(total_files, total_files)
        return pdf_files

    @staticmethod
    def count_pages_in_pdf(filepath):
        try:
            with open(filepath, 'rb') as pdf_file:
                reader = PdfReader(pdf_file)
                if reader.is_encrypted:
                    try:
                        reader.decrypt('')
                    except Exception as e:
                        return (os.path.basename(filepath), None, os.path.dirname(filepath), f"Could not decrypt {filepath}: {e}")
                return (os.path.basename(filepath), len(reader.pages), os.path.dirname(filepath), None)
        except Exception as e:
            return (os.path.basename(filepath), None, os.path.dirname(filepath), f"Error reading {filepath}: {e}")

    def count_pdf_pages(self, pdf_files, progress_callback):
        num_workers = multiprocessing.cpu_count()
        results, errors = [], []
        total_files = len(pdf_files)
        progress_callback(0, total_files)
        with ProcessPoolExecutor(max_workers=num_workers) as executor:
            future_to_pdf = {executor.submit(
                PDFPageCounterApp.count_pages_in_pdf, pdf): pdf for pdf in pdf_files}
            for i, future in enumerate(as_completed(future_to_pdf)):
                if self.stop_event.is_set():
                    executor.shutdown(wait=False)
                    raise Exception("Script stopped by user.")
                result = future.result()
                if result[1] is not None:
                    results.append(result)
                else:
                    errors.append(result)
                if (i + 1) % PROCESSING_PROGRESS_UPDATE_INTERVAL == 0 or (i + 1) == total_files:
                    progress_callback(i + 1, total_files)
        return results, errors

    def save_results_to_excel(self, folder_path, results, errors):
        excel_save_path = os.path.join(folder_path, 'pdf_page_counts.xlsx')
        df_results = pd.DataFrame(results, columns=[
                                  'Filename', 'Number of Pages', 'Directory', 'Error']).drop(columns=['Error'])
        df_errors = pd.DataFrame(
            errors, columns=['Filename', 'Number of Pages', 'Directory', 'Error'])
        with pd.ExcelWriter(excel_save_path, engine='openpyxl') as writer:
            df_results.to_excel(writer, sheet_name='PDF Info', index=False)
            df_errors.to_excel(writer, sheet_name='Errors', index=False)

    def display_success(self, message):
        messagebox.showinfo("Success", message)

    def reset_progress_bars(self):
        self.update_processing_progress(0, 0)
        self.update_search_progress(0, 0)


if __name__ == "__main__":
    root = Tk()
    root.title("PDF Page Counter")
    app = PDFPageCounterApp(root)
    root.mainloop()
