#!/usr/bin/env python
# coding: utf-8

# In[2]:


import os
import zipfile
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import csv
import shutil
import threading
import re
import hashlib
import queue

MAX_FOLDER_NAME_LENGTH = 50  # Max length for each folder name segment
MAX_PATH_LENGTH = 260        # Max length for the entire path

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Rebuild Folder Structure and Extract Files from ZIPs")

        self.zip_files = []
        self.output_folder = ""

        # Progress tracker
        self.total_files = 0
        self.files_processed = 0
        self.files_errors = 0

        # Error processing tracker
        self.total_errors = 0
        self.errors_fixed = 0
        self.errors_failed = 0

        self.sanitized_paths_set = set()
        self.queue = queue.Queue()
        self.create_widgets()
        self.root.after(100, self.process_queue)

    def create_widgets(self):
        instruction_label = tk.Label(self.root, text="Select ZIP files and an output directory")
        instruction_label.pack(pady=10)

        # Zipped folders
        self.zip_label = tk.Label(self.root, text="Select ZIP Folders:")
        self.zip_label.pack()
        self.zip_button = tk.Button(self.root, text="Browse", command=self.select_zip_files)
        self.zip_button.pack()
        self.zip_files_list = tk.Listbox(self.root, width=80)
        self.zip_files_list.pack()

        # Output folder
        self.output_label = tk.Label(self.root, text="Select Output Folder:")
        self.output_label.pack()
        self.output_button = tk.Button(self.root, text="Browse", command=self.select_output_folder)
        self.output_button.pack()
        self.output_folder_label = tk.Label(self.root, text="")
        self.output_folder_label.pack()

        # Start button
        self.start_button = tk.Button(self.root, text="Start Processing", command=self.start_processing)
        self.start_button.pack(pady=10)

        # Progress tracker
        self.progress_label_extracting = tk.Label(self.root, text="")
        self.progress_label_extracting.pack()

        self.progress_label_errors = tk.Label(self.root, text="")
        self.progress_label_errors.pack()

        # Progress bar
        self.progress_bar = ttk.Progressbar(self.root, orient="horizontal", length=400, mode="determinate")
        self.progress_bar.pack(pady=5)

    def select_zip_files(self):
        files = filedialog.askopenfilenames(title="Select ZIP files", filetypes=[("ZIP files", "*.zip")])
        if files:
            self.zip_files = list(files)
            self.zip_files_list.delete(0, tk.END)
            for file in self.zip_files:
                self.zip_files_list.insert(tk.END, file)

    def select_output_folder(self):
        folder = filedialog.askdirectory(title="Select Output Directory")
        if folder:
            self.output_folder = folder
            self.output_folder_label.config(text=self.output_folder)

    def start_processing(self):
        if not self.zip_files:
            messagebox.showerror("Error", "Please select at least one ZIP file.")
            return
        if not self.output_folder:
            messagebox.showerror("Error", "Please select an output folder.")
            return

        # Disable start button
        self.start_button.config(state=tk.DISABLED)

        # Start processing in separate thread
        threading.Thread(target=self.process_zips, daemon=True).start()

    def process_zips(self):
        # 1: Extract folder structure
        self.queue.put(('update_progress_label_extracting', "Extracting files..."))
        self.extract_files()

        # 2: Process errors (if any)
        error_log_file = os.path.join(self.output_folder, 'error_log.csv')
        if os.path.exists(error_log_file):
            self.queue.put(('update_progress_label_errors', "Processing errors..."))
            self.process_errors(error_log_file)

        # Write processing summary to text file
        self.write_processing_summary()

        # Enable start button
        self.queue.put(('enable_start_button', None))

    def extract_files(self):
        errors = []
        processed_entries = 0

        # Calculate total number of files for progress tracking
        total_files = 0
        for zip_file in self.zip_files:
            with zipfile.ZipFile(zip_file, 'r') as zf:
                total_files += len([f for f in zf.namelist() if not f.endswith('/')])

        self.total_files = total_files

        # Log status for each file
        csv_status_file = os.path.join(self.output_folder, 'file_status.csv')
        fieldnames = ['Original File Path', 'Original File Name', 'Sanitized File Name', 'Original Destination Path',
                      'Sanitized Destination Path', 'Moved', 'Reason']

        with open(csv_status_file, 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()

            for zip_file in self.zip_files:
                try:
                    with zipfile.ZipFile(zip_file, 'r') as zf:
                        all_files = zf.namelist()

                        # Filter out folders
                        all_files = [f for f in all_files if not f.endswith('/')]

                        for file in all_files:
                            processed_entries += 1
                            # Update progress
                            progress = processed_entries / total_files * 100
                            self.queue.put(('update_progress', progress))
                            self.queue.put(('update_progress_label_extracting',
                                            f"Extracting files: {processed_entries} of {total_files}"))

                            # Initialize file status dictionary
                            file_status = {
                                'Original File Path': file,
                                'Original File Name': os.path.basename(file),
                                'Sanitized File Name': '',
                                'Original Destination Path': '',
                                'Sanitized Destination Path': '',
                                'Moved': 'False',
                                'Reason': ''
                            }

                            # Remove any leading drive letters and slashes, replace backslashes
                            file_norm = re.sub(r'^[a-zA-Z]:', '', file)
                            file_norm = file_norm.lstrip('/\\')
                            file_norm = file_norm.replace('\\', '/')

                            # Sanitize the path (including trailing spaces)
                            sanitized_file_mapped = self.sanitize_path(file_norm)

                            # Update file status
                            file_status['Sanitized File Name'] = os.path.basename(sanitized_file_mapped)
                            file_status['Original Destination Path'] = os.path.join(self.output_folder, file_norm)
                            file_status['Sanitized Destination Path'] = os.path.join(self.output_folder,
                                                                                   sanitized_file_mapped)

                            self.sanitized_paths_set.add(file_status['Sanitized Destination Path'])

                            # Set destination path
                            dest_path = os.path.normpath(file_status['Sanitized Destination Path'])
                            dest_dir = os.path.dirname(dest_path)

                            # Make sure dest_dir exists
                            if not os.path.exists(dest_dir):
                                os.makedirs(dest_dir, exist_ok=True)

                            # Avoid filename collisions with hashes
                            original_dest_path = dest_path
                            collision_count = 0
                            while dest_path in self.sanitized_paths_set or os.path.exists(dest_path):
                                collision_count += 1
                                filename, ext = os.path.splitext(os.path.basename(dest_path))
                                hash_input = f"{filename}_{collision_count}"
                                short_hash = hashlib.md5(hash_input.encode()).hexdigest()[:8]

                                # Adjust filename length to based on hash and extension
                                max_filename_length = MAX_FOLDER_NAME_LENGTH - len(ext) - len(short_hash) - 1
                                if max_filename_length < 1:
                                    max_filename_length = 1
                                filename = filename[:max_filename_length]

                                # Create new filename
                                filename = f"{filename}_{short_hash}"
                                dest_path = os.path.join(dest_dir, filename + ext)

                            # Update sanitized path and file status if changed
                            if dest_path != original_dest_path:
                                file_status['Sanitized Destination Path'] = dest_path
                                file_status['Sanitized File Name'] = os.path.basename(dest_path)

                            self.sanitized_paths_set.add(dest_path)

                            # Make sure path length is within the limit
                            if len(dest_path) > MAX_PATH_LENGTH:
                                # Shorten the path if needed
                                try:
                                    dest_path = self.shorten_path(dest_path, sanitized_file_mapped)
                                except Exception as e:
                                    errors.append(
                                        {'zip_file': zip_file, 'file': file, 'error_message': str(e)})
                                    self.files_errors += 1
                                    file_status['Moved'] = 'False'
                                    file_status['Reason'] = str(e)
                                    writer.writerow(file_status)
                                    self.update_file_progress()
                                    continue

                            # Extract the file
                            try:
                                with zf.open(file) as source, open(dest_path, 'wb') as target:
                                    shutil.copyfileobj(source, target)
                                self.files_processed += 1
                                file_status['Moved'] = 'True'
                            except Exception as e:
                                errors.append({'zip_file': zip_file, 'file': file, 'error_message': str(e)})
                                self.files_errors += 1
                                file_status['Moved'] = 'False'
                                file_status['Reason'] = str(e)
                                writer.writerow(file_status)
                                self.update_file_progress()
                                continue

                            # Write the file status to CSV
                            writer.writerow(file_status)

                            # Update progress
                            self.update_file_progress()

                except Exception as e:
                    errors.append({'zip_file': zip_file, 'file': '', 'error_message': f"Error processing ZIP file: {str(e)}"})
                    self.files_errors += 1
                    self.update_file_progress()
                    continue

        # Write errors to CSV
        if errors:
            error_log_file = os.path.join(self.output_folder, 'error_log.csv')
            with open(error_log_file, 'w', newline='', encoding='utf-8') as csvfile:
                writer = csv.DictWriter(csvfile, fieldnames=['zip_file', 'file', 'error_message'])
                writer.writeheader()
                for error in errors:
                    writer.writerow(error)

    def sanitize_path(self, file_path):
        # Replace backslashes with slashes
        file_path = file_path.replace('\\', '/')

        # Split into parts
        parts = file_path.split('/')

        sanitized_parts = []
        for i, part in enumerate(parts):
            # Remove leading and trailing spaces
            part = part.strip()

            # Remove invalid characters
            invalid_chars = r'<>:"/\\|?*'
            part = re.sub(f'[{re.escape(invalid_chars)}]', '_', part)

            if i == len(parts) - 1:
                # Preserve the extension
                filename, ext = os.path.splitext(part)

                # Ensure total filename length does not exceed MAX_FOLDER_NAME_LENGTH
                max_filename_length = MAX_FOLDER_NAME_LENGTH - len(ext)
                if max_filename_length < 1:
                    filename = hashlib.md5(filename.encode()).hexdigest()[:8]
                else:
                    filename = filename[:max_filename_length]

                part = filename + ext
            else:
                # Truncate directory names
                part = part[:MAX_FOLDER_NAME_LENGTH]

            sanitized_parts.append(part)

        # Reconstruct the path
        sanitized_path = os.path.join(*sanitized_parts)

        # Ensure total path length does not exceed MAX_PATH_LENGTH
        full_path = os.path.abspath(os.path.join(self.output_folder, sanitized_path))
        if len(full_path) > MAX_PATH_LENGTH:
            # Shorten the path
            try:
                full_path = self.shorten_path(full_path, sanitized_parts)
            except Exception as e:
                raise Exception(f"Cannot shorten path: {sanitized_path}. Error: {e}")

        return os.path.relpath(full_path, self.output_folder)

    def shorten_path(self, full_path, sanitized_parts):
        # Ensure total path length does not exceed MAX_PATH_LENGTH
        max_total_length = MAX_PATH_LENGTH
        output_folder_length = len(os.path.abspath(self.output_folder))
        max_path_length = max_total_length - output_folder_length - 1  # Subtract 1 for the separator

        # If the full path is already within the limit, just return it
        if len(full_path) <= max_total_length:
            return full_path

        # Start shortening file names from the deepest directory going up
        for i in range(len(sanitized_parts)-1, -1, -1):
            part = sanitized_parts[i]
            if i == len(sanitized_parts) - 1:
                # Preserve the extension
                filename, ext = os.path.splitext(part)
                if len(filename) > 8:
                    filename = filename[:8]
                else:
                    filename = filename[:max(1, len(filename) - 1)]
                sanitized_parts[i] = filename + ext
            else:
                # Shorten directory names
                if len(sanitized_parts[i]) > 8:
                    sanitized_parts[i] = sanitized_parts[i][:8]
                else:
                    sanitized_parts[i] = sanitized_parts[i][:max(1, len(sanitized_parts[i]) - 1)]

            # Reconstruct the path and check its length
            new_full_path = os.path.abspath(os.path.join(self.output_folder, *sanitized_parts))
            if len(new_full_path) <= max_total_length:
                return new_full_path

        # If we reach here, we couldn't shorten the path sufficiently
        # As a last resort, we can hash parts of the path
        for i in range(len(sanitized_parts)):
            hashed_part = hashlib.md5(sanitized_parts[i].encode()).hexdigest()[:6]
            sanitized_parts[i] = hashed_part
            new_full_path = os.path.abspath(os.path.join(self.output_folder, *sanitized_parts))
            if len(new_full_path) <= max_total_length:
                return new_full_path

        # If still too long, raise an exception
        raise Exception("Cannot shorten path to acceptable length.")

    def update_file_progress(self):
        # Put progress updates into the queue
        self.queue.put(('update_file_progress', {
            'total_files': self.total_files,
            'files_processed': self.files_processed,
            'files_errors': self.files_errors
        }))

    def process_errors(self, error_log_file):
        # Read the error_log.csv
        with open(error_log_file, 'r', newline='', encoding='utf-8') as csvfile:
            reader = csv.DictReader(csvfile)
            errors = list(reader)

        # Prepare to log fixed and final errors
        fixed_errors = []
        final_errors = []

        total_errors = len(errors)
        self.total_errors = total_errors
        errors_fixed = 0
        errors_failed = 0

        # Open the CSV file to append file statuses
        csv_status_file = os.path.join(self.output_folder, 'file_status.csv')
        with open(csv_status_file, 'a', newline='', encoding='utf-8') as csvfile:
            fieldnames = ['Original File Path', 'Original File Name', 'Sanitized File Name', 'Original Destination Path',
                          'Sanitized Destination Path', 'Moved', 'Reason']
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)

            for idx, error in enumerate(errors, 1):
                zip_file = error['zip_file']
                file = error['file']
                error_message = error['error_message']

                # Skip errors that are not related to files (e.g., zip file errors)
                if not file:
                    final_errors.append(error)
                    errors_failed += 1
                    continue

                # Initialize dictionary
                file_status = {
                    'Original File Path': file,
                    'Original File Name': os.path.basename(file),
                    'Sanitized File Name': '',
                    'Original Destination Path': '',
                    'Sanitized Destination Path': '',
                    'Moved': 'False',
                    'Reason': ''
                }

                try:
                    # Sanitize the path
                    sanitized_file_mapped = self.sanitize_path(file)

                    # Update file status
                    file_status['Sanitized File Name'] = os.path.basename(sanitized_file_mapped)
                    file_status['Original Destination Path'] = os.path.join(self.output_folder, file)
                    file_status['Sanitized Destination Path'] = os.path.join(self.output_folder, sanitized_file_mapped)

                    # Attempt to extract the file from zip_file
                    dest_path = file_status['Sanitized Destination Path']
                    dest_dir = os.path.dirname(dest_path)

                    # Ensure dest_dir exists
                    if not os.path.exists(dest_dir):
                        os.makedirs(dest_dir, exist_ok=True)

                    with zipfile.ZipFile(zip_file, 'r') as zf:
                        with zf.open(file) as source, open(dest_path, 'wb') as target:
                            shutil.copyfileobj(source, target)

                    errors_fixed += 1
                    file_status['Moved'] = 'True'

                    # Log fixed error
                    fixed_errors.append({
                        'zip_file': zip_file,
                        'original_file': file,
                        'sanitized_file': sanitized_file_mapped,
                        'status': 'Fixed'
                    })

                    # Write the file status to CSV
                    writer.writerow(file_status)

                except Exception as e:
                    errors_failed += 1
                    # Log final error
                    final_errors.append({
                        'zip_file': zip_file,
                        'file': file,
                        'error_message': str(e)
                    })

                    # Update file status
                    file_status['Moved'] = 'False'
                    file_status['Reason'] = str(e)

                    # Write file status to CSV
                    writer.writerow(file_status)
                    continue

                # Update progress
                self.queue.put(('update_progress_label_errors', f"Processing errors: {idx}/{total_errors}, "
                                                               f"Fixed: {errors_fixed}, Failed: {errors_failed}"))

        # Update class variables
        self.errors_fixed = errors_fixed
        self.errors_failed = errors_failed

        # Write fixed_errors.csv
        if fixed_errors:
            fixed_errors_file = os.path.join(self.output_folder, 'fixed_errors.csv')
            with open(fixed_errors_file, 'w', newline='', encoding='utf-8') as csvfile:
                writer = csv.DictWriter(csvfile, fieldnames=['zip_file', 'original_file', 'sanitized_file', 'status'])
                writer.writeheader()
                for item in fixed_errors:
                    writer.writerow(item)

        # Write final_errors.csv
        if final_errors:
            final_errors_file = os.path.join(self.output_folder, 'final_errors.csv')
            with open(final_errors_file, 'w', newline='', encoding='utf-8') as csvfile:
                writer = csv.DictWriter(csvfile, fieldnames=['zip_file', 'file', 'error_message'])
                writer.writeheader()
                for item in final_errors:
                    writer.writerow(item)

        # Update final progress
        self.queue.put(('update_progress_label_errors',
                        f"Error processing complete. Total errors: {total_errors}, "
                        f"Fixed: {errors_fixed}, Failed: {errors_failed}"))

    def write_processing_summary(self):
        summary_file = os.path.join(self.output_folder, 'processing_summary.txt')
        with open(summary_file, 'w', encoding='utf-8') as file:
            file.write(f"Total files processed: {self.total_files}\n")
            file.write(f"Files successfully extracted: {self.files_processed}\n")
            file.write(f"Files with errors: {self.files_errors}\n")
            if self.total_errors > 0:
                file.write(f"\nError Processing Summary:\n")
                file.write(f"Total errors: {self.total_errors}\n")
                file.write(f"Errors fixed: {self.errors_fixed}\n")
                file.write(f"Errors failed: {self.errors_failed}\n")

    def process_queue(self):
        try:
            while True:
                msg = self.queue.get_nowait()
                if msg[0] == 'update_progress_label_extracting':
                    self.progress_label_extracting.config(text=msg[1])
                elif msg[0] == 'update_progress_label_errors':
                    self.progress_label_errors.config(text=msg[1])
                elif msg[0] == 'update_progress':
                    self.progress_bar["value"] = msg[1]
                    self.progress_bar.update_idletasks()
                elif msg[0] == 'enable_start_button':
                    self.start_button.config(state=tk.NORMAL)
                elif msg[0] == 'update_file_progress':
                    data = msg[1]
                    self.progress_label_extracting.config(text=f"Extracting files: Total files: {data['total_files']}, "
                                                               f"Processed: {data['files_processed']}, "
                                                               f"Errors: {data['files_errors']}")
        except queue.Empty:
            pass
        self.root.after(100, self.process_queue)

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()

