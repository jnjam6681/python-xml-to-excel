import xml.etree.ElementTree as ET
import pandas as pd
import os
import glob
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import List

class XMLToExcelConverter:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("XML Converter")
        self.window.geometry("600x700")  # Increased initial height
        self.window.minsize(600, 700)    # Set minimum window size
        self.files: List[str] = []
        self.setup_ui()

    def setup_ui(self):
        # Main container
        main_container = tk.Frame(self.window)
        main_container.pack(fill="both", expand=True)

        # Scrollable content frame
        content_frame = tk.Frame(main_container)
        content_frame.pack(fill="both", expand=True, padx=20, pady=(20,60))  # Added bottom padding for button

        # Title (in content frame)
        title = tk.Label(content_frame, text="XML Converter", font=("Arial", 24))
        title.pack(pady=(0,10))

        subtitle = tk.Label(content_frame, text="อัปโหลดไฟล์ XML และแปลงเป็น Excel", font=("Arial", 14))
        subtitle.pack(pady=(0,10))

        # Upload frame
        upload_frame = tk.Frame(content_frame, padx=20, pady=15)  # Removed bg color
        upload_frame.pack(fill="x")

        # Browse button with enhanced visibility
        browse_btn = ttk.Button(
            upload_frame, 
            text="เลือกไฟล์ XML", 
            command=self.browse_files,
            style='Bold.TButton'
        )
        browse_btn.pack(pady=5)

        # File type hint
        limit_label = tk.Label(
            upload_frame, 
            text="รองรับไฟล์ XML ขนาดไม่เกิน 200MB", 
            font=("Arial", 10, "bold"),
            fg="gray"
        )
        limit_label.pack()

        # Files list frame with scrollbar
        files_container = tk.Frame(content_frame)
        files_container.pack(fill="x", pady=5)
        
        # Create canvas and scrollbar for files
        self.files_canvas = tk.Canvas(files_container, height=80)  # Reduced height
        files_scrollbar = ttk.Scrollbar(files_container, orient="vertical", command=self.files_canvas.yview)
        
        # Configure files frame
        self.files_frame = tk.Frame(self.files_canvas)
        self.files_frame.bind("<Configure>", lambda e: self.files_canvas.configure(scrollregion=self.files_canvas.bbox("all")))
        
        # Add files frame to canvas
        self.files_canvas.create_window((0, 0), window=self.files_frame, anchor="nw", width=self.files_canvas.winfo_reqwidth())
        self.files_canvas.configure(yscrollcommand=files_scrollbar.set)
        
        # Pack canvas and scrollbar
        self.files_canvas.pack(side="left", fill="both", expand=True)
        files_scrollbar.pack(side="right", fill="y")

        # Preview section
        preview_label = tk.Label(content_frame, text="ตัวอย่างข้อมูล:", font=("Arial", 12))
        preview_label.pack(anchor="w", pady=(10,5))

        # Table frame
        table_frame = tk.Frame(content_frame)
        table_frame.pack(fill="both", expand=True)
        self.setup_preview_table(table_frame)

        # Fixed button frame at bottom of window
        button_frame = tk.Frame(self.window, height=50)  # Removed bg color
        button_frame.pack(side="bottom", fill="x")
        button_frame.pack_propagate(False)  # Prevent frame from shrinking

        # Convert button
        convert_btn = ttk.Button(button_frame, text="ดาวน์โหลดไฟล์ Excel", command=self.convert_files)
        convert_btn.pack(side="right", padx=20, pady=10)

    def add_file_to_list(self, filepath: str):
        file_frame = tk.Frame(self.files_frame)
        file_frame.pack(fill="x", pady=1)  # Minimal padding

        # File icon and name
        file_label = tk.Label(file_frame, text=os.path.basename(filepath), font=("Arial", 10))  # Smaller font
        file_label.pack(side="left")

        # File size
        size = os.path.getsize(filepath) / 1024
        size_label = tk.Label(file_frame, text=f"{size:.1f}KB", font=("Arial", 10))
        size_label.pack(side="left", padx=5)  # Reduced padding

        # Remove button
        remove_btn = ttk.Button(file_frame, text="×", width=2,
                              command=lambda: self.remove_file(filepath, file_frame))
        remove_btn.pack(side="right")

    def setup_preview_table(self, parent_frame):
        columns = ("UserName", "ServiceTag", "Operating System")
        
        # Create scrollbar
        scrollbar = ttk.Scrollbar(parent_frame)
        scrollbar.pack(side="right", fill="y")

        # Create treeview with scrollbar
        self.tree = ttk.Treeview(parent_frame, columns=columns, show="headings",
                                yscrollcommand=scrollbar.set)
        
        # Configure scrollbar
        scrollbar.config(command=self.tree.yview)
        
        # Set column headings
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=150)

        self.tree.pack(side="left", fill="both", expand=True)

    def add_file_to_list(self, filepath: str):
        file_frame = tk.Frame(self.files_frame)
        file_frame.pack(fill="x", pady=1)

        # File icon and name
        file_label = tk.Label(file_frame, text=os.path.basename(filepath))
        file_label.pack(side="left")

        # File size
        size = os.path.getsize(filepath) / 1024  # Convert to KB
        size_label = tk.Label(file_frame, text=f"{size:.1f}KB")
        size_label.pack(side="left", padx=10)

        # Remove button
        remove_btn = ttk.Button(file_frame, text="×", width=3,
                              command=lambda: self.remove_file(filepath, file_frame))
        remove_btn.pack(side="right")
        
        # Update canvas scrollregion
        self.files_canvas.configure(scrollregion=self.files_canvas.bbox("all"))

    def remove_file(self, filepath: str, frame: tk.Frame):
        self.files.remove(filepath)
        frame.destroy()
        self.update_preview()

    def browse_files(self):
        filepaths = filedialog.askopenfilenames(
            title="Select XML files",
            filetypes=[("XML files", "*.xml")],
            initialdir=os.getcwd()
        )
        for filepath in filepaths:
            if filepath not in self.files:
                self.files.append(filepath)
                self.add_file_to_list(filepath)
        self.update_preview()

    def update_preview(self):
        # Clear existing items
        for item in self.tree.get_children():
            self.tree.delete(item)

        # Add new preview data
        for filepath in self.files:
            try:
                tree = ET.parse(filepath)
                root = tree.getroot()
                hardware_info = extract_hardware_info(root)
                self.tree.insert("", "end", values=(
                    hardware_info["User Name"],
                    hardware_info["Service Tag"],
                    hardware_info["Operating System"]
                ))
            except Exception as e:
                print(f"Error previewing {filepath}: {str(e)}")

    def convert_files(self):
        if not self.files:
            messagebox.showwarning("Warning", "Please select XML files first!")
            return

        # Ask user for save location
        output_file = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile="system_inventory.xlsx",
            title="Save Excel File"
        )

        if not output_file:  # User cancelled
            return

        all_data = []
        for file_path in self.files:
            try:
                data_rows = process_xml_file(file_path)
                all_data.extend(data_rows)
            except Exception as e:
                messagebox.showerror("Error", f"Error processing {os.path.basename(file_path)}: {str(e)}")

        if all_data:
            df = pd.DataFrame(all_data)
            df.to_excel(output_file, sheet_name="System Information", index=False)
            messagebox.showinfo("Success", f"Data successfully saved to: {output_file}")

    def run(self):
        self.window.mainloop()

# Keep existing helper functions
def extract_hardware_info(root):
    hardware = root.find("Hardware_Info")
    computer = hardware.find(".//Computer")
    os_info = hardware.find(".//OperatingSystem")
    windows_key = root.find(".//WindowsKey")
    
    return {
        "Computer Name": computer.get("Name"),
        "User Name": computer.get("UserName"),
        "Service Tag": computer.get("ServiceTag"),
        "Operating System": os_info.get("Name"),
        "Windows Product ID": windows_key.get("ProductID"),
        "Windows Product Key": windows_key.get("ProductKey"),
    }

def process_xml_file(xml_path):
    tree = ET.parse(xml_path)
    root = tree.getroot()
    hardware_info = extract_hardware_info(root)
    software_list = root.findall(".//InstalledProgramsList/*")
    
    data_rows = []
    for software in software_list:
        row_data = hardware_info.copy()
        row_data["Installed Software"] = software.get("Name")
        data_rows.append(row_data)
    
    return data_rows

if __name__ == "__main__":
    app = XMLToExcelConverter()
    app.run()
