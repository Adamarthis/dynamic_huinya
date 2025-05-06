import tkinter as tk
from tkinter import filedialog, messagebox
import json
import random
import fitz  # PyMuPDF
import io
import os
import traceback
import math

# --- Matplotlib setup for LaTeX ---
# import matplotlib
# matplotlib.use('Agg')  # Use non-interactive backend
# import matplotlib.pyplot as plt
# from matplotlib import mathtext

# Ensure mathtext uses LaTeX (REMOVED)
# matplotlib.rcParams['mathtext.fontset'] = 'cm'
# matplotlib.rcParams['mathtext.rm'] = 'serif'
# --- End Matplotlib setup ---

# --- ReportLab setup ---
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Image, Paragraph, Spacer
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
# Need fitz here
# --- End ReportLab setup ---

# --- Configuration ---
TASKS_JSON_PATH = "tasks.json"
SOURCE_PDF_PATH = "dynamic_sample.pdf"

# NEW Configuration: Map page numbers (1-based) to task details
# Structure: page_num: (task_key, needed_count, structure_info, identifier_text_to_find)
PAGE_TASK_MAP = {
    # Page: (TaskKey, NeededCount, StructureInfo, IdentifierText)
    7:  ("LAB1", 30, {"type": "3_objects"}, "Таблиця 1.1 – Варіанти завдань до лабораторної роботи № 1"),
    22: ("LAB3", 30, {"type": "description"}, "Таблиця 3.1 – Варіанти завдань до лабораторної роботи № 3"),
    27: ("LAB4_TASK1", 14, {"type": "description"}, "Таблиця 4.1 – Варіанти для завдання 1"),
    29: ("LAB4_TASK2", 15, {"type": "description"}, "Таблиця 4.2 – Варіанти для завдання 2"),
    31: ("LAB4_TASK3", 16, {"type": "description"}, "Таблиця 4.3 – Варіанти для завдання 3"),
    44: ("LAB5_TASK1", 19, {"type": "description"}, "Таблиця 5.1 – Варіанти для завдання 1"),
    46: ("LAB5_TASK2", 18, {"type": "description"}, "Таблиця 5.2 – Варіанти для завдання 2"),
    58: ("LAB6", 20, {"type": "pair_description"}, "Таблиця 6.1 – Індивідуальні варіанти завдань"),
    73: ("LAB7_TASK1", 32, {"type": "description"}, "Таблиця 7.1 – Варіанти до завдання 1"),
    86: ("LAB9_TASK1", 30, {"type": "description"}, "Таблиця 8.1 – Варіанти до завдання 1"),
    88: ("LAB9_TASK2", 30, {"type": "description"}, "Таблиця 8.2 – Варіанти до завдання 2"),
    102: ("LAB10_TASK1", 27, {"type": "description"}, "Таблиця 10.1 – Варіанти для першого завдання лабораторної роботи"),
    103: ("LAB10_TASK2", 23, {"type": "description"}, "Таблиця 10.2 – Варіанти для другого завдання"), # Inserted on p103
    107: ("LAB11", 20, {"type": "description"}, "Таблиця 11.1 – Індивідуальні варіанти завдань"),
    # Lab 13 is missing from the user's list? Assuming skip or was error.
    # 107: ("LAB13", 20, {"type": "profession_pairs"}, "Таблиця 13.1 – Індивідуальні варіанти завдань"),
    137: ("LAB16", 20, {"type": "function_pair"}, "у другому стовпці – задані функції x(t) та y(t)."), # Note: No latex=True needed
}

# --- Helper Functions ---

# REMOVED render_latex_to_image function
# def render_latex_to_image(latex_str, dpi=150):
#    ... (Removed) ...


# --- Main Application Class ---

class PdfGeneratorApp:
    def __init__(self, master):
        self.master = master
        master.title("Dynamic PDF Generator")
        master.geometry("400x200") # Adjusted size

        # --- Font Setup --- Restore TTF attempt with strict verification
        font_regular_name = 'Times-Roman' # Default fallback
        font_bold_name = 'Times-Bold'   # Default fallback
        registered_name_reg = 'MyDejaVuSans' # Use a unique name
        registered_name_bold = 'MyDejaVuSans-Bold'
        registration_success = False
        try:
            # Corrected paths based on user confirmation
            user_profile = os.environ.get('USERPROFILE')
            if user_profile:
                font_dir = os.path.join(user_profile, 'AppData', 'Local', 'Microsoft', 'Windows', 'Fonts')
                font_path_reg = os.path.join(font_dir, "DejaVuSans.ttf")
                font_path_bold = os.path.join(font_dir, "DejaVuSans-Bold.ttf")
            else:
                # Fallback or raise error if user profile path can't be determined
                print("!!! WARNING: Cannot determine user profile path. Falling back to C:\\Windows\\Fonts.")
                font_path_reg = "C:\\Windows\\Fonts\\DejaVuSans.ttf"
                font_path_bold = "C:\\Windows\\Fonts\\DejaVuSans-Bold.ttf"

            if os.path.exists(font_path_reg) and os.path.exists(font_path_bold):
                 print(f"Font path found: {font_path_reg}")
                 pdfmetrics.registerFont(TTFont(registered_name_reg, font_path_reg))
                 print(f"Attempting to register TTF: {font_path_bold} as {registered_name_bold}")
                 pdfmetrics.registerFont(TTFont(registered_name_bold, font_path_bold))

                 # --- Strict Verification --- >
                 try:
                     # Try getting the font - this fails if registration didn't truly work
                     pdfmetrics.getFont(registered_name_reg)
                     pdfmetrics.getFont(registered_name_bold)
                     # If we get here, registration likely succeeded
                     font_regular_name = registered_name_reg
                     # Use the registered bold font name
                     font_bold_name = registered_name_bold
                     registration_success = True
                     print(f"Successfully registered and VERIFIED TTF: {font_regular_name}, {font_bold_name}")
                 except pdfmetrics.FontNotFoundError:
                     print(f"!!! CRITICAL WARNING: TTF fonts exist and pdfmetrics.registerFont called, but getFont('{registered_name_reg}') FAILED. Check permissions or font integrity.")
                     print(f"    Available fonts after attempt: {pdfmetrics.getRegisteredFontNames()}")
                     print(f"    Falling back to: Times-Roman/Times-Bold (Encoding WILL likely fail)")
                     font_regular_name = 'Times-Roman' # Explicitly fall back regular too
                     font_bold_name = 'Times-Bold'
                 # < --- Strict Verification ---
            else:
                 print(f"!!! WARNING: DejaVuSans TTF files not found at specified paths. Using fallback Times-Roman/Times-Bold (Encoding WILL likely fail).")
                 font_regular_name = 'Times-Roman'
                 font_bold_name = 'Times-Bold'
        except Exception as e:
            print(f"!!! ERROR during font registration: {e}. Using fallback Times-Roman/Times-Bold (Encoding WILL likely fail).")
            font_regular_name = 'Times-Roman'
            font_bold_name = 'Times-Bold'

        # Store final font names
        self.font_regular = font_regular_name
        self.font_bold = font_bold_name # Use the determined bold font (DejaVu or Times)
        # Create styles using the determined fonts
        self.styles = self.create_styles(self.font_regular, self.font_bold)
        # --- End Font Setup ---

        self.label = tk.Label(master, text="Click the button to generate the dynamic PDF.")
        self.label.pack(pady=10)

        self.generate_button = tk.Button(master, text="Generate PDF", command=self.run_generation)
        self.generate_button.pack(pady=10)

        self.status_label = tk.Label(master, text="")
        self.status_label.pack(pady=10)

        # Load tasks on initialization
        self.tasks_data = self.load_tasks()
        if not self.tasks_data:
            self.generate_button.config(state=tk.DISABLED)
            self.status_label.config(text="Error loading tasks.json. Cannot proceed.")

    def create_reportlab_table(self, data, col_widths=None, style_commands=None, row_heights=None):
        """Creates a ReportLab Table object with basic styling."""
        table = Table(data, colWidths=col_widths, rowHeights=row_heights)
        # Base font uses self.font_regular (hopefully MyDejaVuSans)
        # Bold font uses self.font_bold (hopefully MyDejaVuSans-Bold or Times-Bold fallback)
        base_style = [
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, -1), self.font_regular),
            # Header style
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), self.font_bold),
            ('FONTSIZE', (0, 0), (-1, 0), 19), # Larger header font (e.g., 19pt)
            ('BOTTOMPADDING', (0, 0), (-1, 0), 5*mm),
            # First column style (variant number)
            ('ALIGN', (0, 1), (0, -1), 'CENTER'),
            # Use determined bold font for variant numbers
            ('FONTNAME', (0, 1), (0, -1), self.font_bold),
        ]
        if style_commands:
            base_style.extend(style_commands)
        table.setStyle(TableStyle(base_style))
        return table

    def create_styles(self, font_name_regular, font_name_bold):
        """Creates ParagraphStyles, inheriting regular font where possible."""
        # Use the font names determined in __init__
        styles = getSampleStyleSheet()
        base_size = 14 # Keep this for potential non-table text if needed
        table_font_size = 18 # Increase Paragraph font size for table cells to 18pt
        table_leading = table_font_size * 1.2 # Set leading based on font size (e.g., 21.6)

        # Define base styles using the determined regular font
        styles.add(ParagraphStyle(name='Normal_UA', parent=styles['Normal'], fontName=font_name_regular, fontSize=base_size))
        styles.add(ParagraphStyle(name='BodyText_UA', parent=styles['BodyText'], fontName=font_name_regular, fontSize=base_size))
        styles.add(ParagraphStyle(name='Italic_UA', parent=styles['Italic'], fontName=font_name_regular, fontSize=base_size))

        # Headings using determined bold font
        styles.add(ParagraphStyle(name='Heading1_UA', parent=styles['h1'], fontName=font_name_bold, fontSize=base_size+4))
        styles.add(ParagraphStyle(name='Heading2_UA', parent=styles['h2'], fontName=font_name_bold, fontSize=table_font_size+1)) # e.g., 19pt for table headers

        # Table cell paragraph styles - add explicit leading
        # Inherit regular font from Normal_UA
        styles.add(ParagraphStyle(name='TableCell', parent=styles['Normal_UA'], fontSize=table_font_size, alignment=1, leading=table_leading))
        # Use determined bold font
        styles.add(ParagraphStyle(name='TableCellBold', parent=styles['Normal_UA'], fontName=font_name_bold, fontSize=table_font_size, alignment=1, leading=table_leading))
        # Inherit regular font from Normal_UA
        styles.add(ParagraphStyle(name='TableCellLatex', parent=styles['Normal_UA'], fontSize=table_font_size, alignment=1, leading=table_leading))
        # Inherit regular font from Normal_UA
        styles.add(ParagraphStyle(name='TableCellLeft', parent=styles['Normal_UA'], fontSize=table_font_size, alignment=0, leading=table_leading))
        return styles

    def load_tasks(self):
        """Loads the tasks from the JSON file."""
        try:
            with open(TASKS_JSON_PATH, 'r', encoding='utf-8') as f:
                return json.load(f)
        except FileNotFoundError:
            messagebox.showerror("Error", f"Tasks file not found: {TASKS_JSON_PATH}")
            return None
        except json.JSONDecodeError as e:
            messagebox.showerror("Error", f"Error decoding JSON from {TASKS_JSON_PATH}:\n{e}")
            return None
        except Exception as e:
            messagebox.showerror("Error", f"An unexpected error occurred loading tasks: {e}")
            return None

    def select_unique_tasks(self, all_tasks_data):
        """Selects unique tasks for all configured pages."""
        selected_tasks_map = {} # Key: page_num, Value: list of selected tasks
        # Use a single pool for overall uniqueness check
        overall_used_tasks = set() # Key: (task_key, index)

        # Prepare available indices for each task key
        available_tasks = {}
        for key, tasks in all_tasks_data.items():
            available_tasks[key] = list(range(len(tasks)))

        # Iterate through the page configuration
        for page_num, (task_key, needed_count, structure_info, _) in PAGE_TASK_MAP.items():
            if task_key not in available_tasks:
                print(f"Warning: Task key '{task_key}' not found in {TASKS_JSON_PATH} for page {page_num}. Skipping.")
                selected_tasks_map[page_num] = []
                continue

            current_selection_indices = []
            potential_indices = available_tasks[task_key][:] # Use a copy
            random.shuffle(potential_indices)

            count = 0
            actual_needed = needed_count
            # Adjust needed count based on structure (e.g., Lab 1 needs 3 per variant)
            if structure_info.get('type') == '3_objects': actual_needed = needed_count * 3
            elif structure_info.get('type') in ['object_pairs', 'pair_description', 'profession_pairs']: actual_needed = needed_count * 2
            #elif structure_info.get('type') == 'function_pair': actual_needed = needed_count # Already 1 per row
            #elif structure_info.get('type') == 'description': actual_needed = needed_count # Already 1 per row

            for index in potential_indices:
                task_tuple = (task_key, index)
                if task_tuple not in overall_used_tasks:
                    current_selection_indices.append(index)
                    overall_used_tasks.add(task_tuple)
                    count += 1
                    if count >= actual_needed:
                        break

            if count < actual_needed:
                 print(f"Warning: Could only select {count}/{actual_needed} unique tasks for page {page_num} ('{task_key}'). Not enough unique tasks available overall.")
                 # Proceeding with the selection made

            # Get the actual task strings/data
            selected_tasks = [all_tasks_data[task_key][i] for i in current_selection_indices]
            selected_tasks_map[page_num] = selected_tasks # Store tasks against page number

        return selected_tasks_map

    def run_generation(self):
        """Handles the button click event to generate the PDF."""
        if not self.tasks_data:
            messagebox.showerror("Error", "Task data is not loaded. Cannot generate PDF.")
            return

        output_pdf_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF Documents", "*.pdf")],
            title="Save Generated PDF As"
        )
        if not output_pdf_path:
            self.status_label.config(text="PDF generation cancelled.")
            return

        self.status_label.config(text="Generating PDF... Please wait.")
        self.master.update()

        output_doc = None
        source_doc = None
        table_pdf = None

        # --- Select ALL tasks needed across all pages first ---
        try:
            # Pass the full tasks_data to the selection function
            selected_tasks_for_pages = self.select_unique_tasks(self.tasks_data)
        except KeyError as e:
            self.status_label.config(text="Error: Task key mismatch.")
            messagebox.showerror("Config Error", f"Task key '{e}' found in PAGE_TASK_MAP but not in {TASKS_JSON_PATH}. Check configuration.")
            return
        except Exception as e:
             self.status_label.config(text="Error during task selection.")
             tb_str = traceback.format_exc()
             messagebox.showerror("Error", f"An unexpected error occurred selecting tasks:\\n{e}\\n\\nTraceback:\\n{tb_str}")
             return

        # --- PDF Generation ---
        try:
            if not os.path.exists(SOURCE_PDF_PATH):
                 raise FileNotFoundError(f"Source PDF not found: {SOURCE_PDF_PATH}")
            source_doc = fitz.open(SOURCE_PDF_PATH)
            output_doc = fitz.open() # Create a new empty PDF for output

            # --- Iterate through source pages, copy, and insert tables ---
            for page_num in range(len(source_doc)):
                page = source_doc.load_page(page_num)
                output_page = output_doc.new_page(width=page.rect.width, height=page.rect.height)
                # Copy the entire source page content first
                output_page.show_pdf_page(output_page.rect, source_doc, page_num)

                # Check if a table needs to be inserted on this page (using 1-based index)
                current_page_1_based = page_num + 1
                if current_page_1_based in PAGE_TASK_MAP:
                    task_key, needed_count, structure_info, identifier_text = PAGE_TASK_MAP[current_page_1_based]
                    selected_tasks = selected_tasks_for_pages.get(current_page_1_based)
                    if not selected_tasks: continue
                    text_instances = output_page.search_for(identifier_text, quads=True)
                    if not text_instances: continue
                    first_instance_quad = text_instances[0]
                    identifier_rect = first_instance_quad.rect
                    insert_y_pos = identifier_rect.y1 + 5
                    table_x0 = page.rect.x0 + 20*mm
                    table_x1 = page.rect.x1 - 20*mm

                    # --- Prepare & Draw New Table ---
                    try:
                        table_data = []
                        col_widths = None
                        table_style_cmds = []
                        row_heights = None

                        # --- Build table_data using correct styles and TEST CHAR ---
                        if structure_info['type'] == '3_objects':
                             # Replace № with No.
                             headers = [ "No.", "Об'єкт 1", "Об'єкт 2", "Об'єкт 3" ]
                             table_data.append([Paragraph(h, self.styles['Heading2_UA']) for h in headers])
                             num_variants = needed_count
                             tasks_per_variant = 3
                             current_task_idx = 0
                             for i in range(num_variants):
                                 row = [Paragraph(str(i + 1), self.styles['TableCellBold'])]
                                 for _ in range(tasks_per_variant):
                                     if current_task_idx < len(selected_tasks):
                                         task_text_raw = selected_tasks[current_task_idx]
                                         try: task_text = task_text_raw.split("Об'єкт ")[-1].split(" за допомогою")[0]
                                         except: task_text = task_text_raw
                                         row.append(Paragraph(task_text, self.styles['TableCell']))
                                         current_task_idx += 1
                                     else:
                                         row.append(Paragraph("-", self.styles['TableCell']))
                                 table_data.append(row)
                             col_widths = [15*mm, 51*mm, 52*mm, 52*mm] # Total 170mm
                        elif structure_info['type'] == 'function_pair':
                             # Replace № with No.
                             headers = ['No.', 'Функції x(t) та y(t)']
                             table_data.append([Paragraph(h, self.styles['Heading2_UA']) for h in headers])
                             # Keep function column wide
                             col_widths = [15*mm, 150*mm]
                             for i, func_pair_str in enumerate(selected_tasks):
                                 try:
                                     if isinstance(func_pair_str, (list, tuple)) and len(func_pair_str) >= 2:
                                         formatted_text = f"{func_pair_str[0]}<br/>{func_pair_str[1]}"
                                     elif isinstance(func_pair_str, str):
                                         formatted_text = func_pair_str.replace('\n', '<br/>') # Basic newline replace
                                         if formatted_text.startswith("['") and formatted_text.endswith("']"):
                                             import ast
                                             try:
                                                 pair_list = ast.literal_eval(formatted_text)
                                                 if isinstance(pair_list, list) and len(pair_list) == 2:
                                                     formatted_text = f"{pair_list[0]}<br/>{pair_list[1]}"
                                                 else:
                                                     formatted_text = func_pair_str.replace("\n", "<br/>").replace("['", "").replace("']", "")
                                             except (ValueError, SyntaxError):
                                                 formatted_text = func_pair_str.replace("\n", "<br/>")
                                     else:
                                         formatted_text = str(func_pair_str) # Fallback
                                     cell_content = Paragraph(formatted_text, self.styles['TableCellLeft'])
                                     table_data.append([Paragraph(str(i + 1), self.styles['TableCellBold']), cell_content])
                                 except Exception as format_e:
                                     print(f"Error formatting function pair {i+1}: {func_pair_str}. Error: {format_e}")
                                     table_data.append([Paragraph(str(i + 1), self.styles['TableCellBold']), Paragraph("Formatting Error", self.styles['TableCellLeft'])])
                             table_style_cmds.append(('ALIGN', (1, 1), (1, -1), 'LEFT'))
                             table_style_cmds.append(('VALIGN', (1, 1), (1, -1), 'TOP'))
                        elif structure_info['type'] == 'pair_description':
                             # Replace № with No.
                             headers = ["No.", "Параметр", "Опис"]
                             table_data.append([Paragraph(h, self.styles['Heading2_UA']) for h in headers])
                             # Increased description column width slightly
                             col_widths = [15*mm, 70*mm, 85*mm] # Total 170mm
                             num_variants = needed_count
                             current_task_idx = 0
                             for i in range(num_variants):
                                 if current_task_idx + 1 < len(selected_tasks):
                                     param = selected_tasks[current_task_idx]
                                     desc = selected_tasks[current_task_idx + 1]
                                     row = [Paragraph(str(i + 1), self.styles['TableCellBold']),
                                            Paragraph(str(param), self.styles['TableCellLeft']),
                                            Paragraph(str(desc), self.styles['TableCellLeft'])]
                                     current_task_idx += 2
                                 else:
                                     row = [Paragraph(str(i + 1), self.styles['TableCellBold']), Paragraph("-", self.styles['TableCellLeft']), Paragraph("-", self.styles['TableCellLeft'])]
                                 table_data.append(row)
                             table_style_cmds.append(('ALIGN', (1, 1), (-1, -1), 'LEFT'))
                        elif structure_info['type'] == 'profession_pairs':
                             # Replace № with No.
                            headers = [ "No.", "Професія 1", "No.", "Професія 2" ]
                            table_data.append([Paragraph(h, self.styles['Heading2_UA']) for h in headers])
                            # Keep column widths (already at 170mm total)
                            col_widths = [15*mm, 70*mm, 15*mm, 70*mm]
                            table_style_cmds.append(('ALIGN', (1, 1), (1, -1), 'LEFT'))
                            table_style_cmds.append(('ALIGN', (3, 1), (3, -1), 'LEFT'))
                        elif structure_info['type'] == 'description':
                             # Replace № with No.
                            headers = ['No.', 'Завдання']
                            table_data.append([Paragraph(h, self.styles['Heading2_UA']) for h in headers])
                            # Keep description column wide
                            col_widths = [15*mm, 150*mm]
                            table_style_cmds.append(('ALIGN', (1, 1), (1, -1), 'LEFT'))
                        else:
                            print(f"Warning: Unhandled table structure type '{structure_info.get('type')}' for page {current_page_1_based}")
                            table_data = [[Paragraph("Error: Unhandled Table Type", self.styles['Normal_UA'])]]
                            col_widths = [170*mm]

                        # --- Draw Table using ReportLab ---
                        if not table_data:
                             print(f"Warning: No table data generated for page {current_page_1_based}, skipping draw.")
                             continue

                        # Create table in memory
                        temp_buffer = io.BytesIO()
                        # Estimate page size needed for the table - use wide width and large height
                        est_page_width = A4[0] # Use A4 width
                        est_page_height = A4[1] * 3 # Assume table won't exceed 3 pages tall
                        temp_doc = SimpleDocTemplate(temp_buffer, pagesize=(est_page_width, est_page_height),
                                                     leftMargin=0, rightMargin=0, # Margins handled by placement rect
                                                     topMargin=0, bottomMargin=0)
                        story = [self.create_reportlab_table(table_data, col_widths, table_style_cmds, row_heights)]
                        temp_doc.build(story)
                        temp_buffer.seek(0)

                        # Open the generated table PDF
                        table_pdf = fitz.open("pdf", temp_buffer.read())
                        if len(table_pdf) > 0:
                            # Use the first page of the generated table
                            table_page = table_pdf.load_page(0)
                            # Calculate target rectangle on the output page
                            target_width = table_x1 - table_x0
                            # Try to use actual table height, but cap it to avoid going off page
                            table_height = table_page.rect.height
                            max_height = output_page.rect.height - insert_y_pos - 10 # Leave 10 points margin at bottom
                            target_height = min(table_height, max_height)

                            if table_height > max_height:
                                print(f"Warning: Table for page {current_page_1_based} is too tall ({table_height:.1f} pts) for available space ({max_height:.1f} pts). It will be clipped.")

                            target_rect = fitz.Rect(table_x0, insert_y_pos, table_x0 + target_width, insert_y_pos + target_height)
                            # Draw the table onto the output page
                            output_page.show_pdf_page(target_rect, table_pdf, 0) # Use page 0 of table_pdf
                        else:
                             print(f"Warning: Generated table PDF for page {current_page_1_based} has no pages.")

                        if table_pdf: table_pdf.close()
                        table_pdf = None
                        temp_buffer.close()

                    except Exception as e: # Catch drawing errors
                        self.status_label.config(text=f"Error drawing table: {identifier_text[:30]}...")
                        tb_str = traceback.format_exc()
                        print(f"ERROR generating/drawing table for page {current_page_1_based}: {e}\n{tb_str}")
                        messagebox.showwarning("Table Generation Warning",
                                               f"Could not generate/draw table for page {current_page_1_based}.\\nIdentifier: '{identifier_text}'\\nError: {e}\\n\\nSkipping this table.")
                        if 'table_pdf' in locals() and table_pdf: table_pdf.close(); table_pdf = None
                        if 'temp_buffer' in locals() and not temp_buffer.closed: temp_buffer.close()
                        continue # Continue to next page

                # No redaction needed or applied here
                output_page.clean_contents() # Clean page contents

            # --- Finalize Output PDF ---
            if len(output_doc) > 0:
                 output_doc.save(output_pdf_path, garbage=4, deflate=True, clean=True)
                 self.status_label.config(text=f"PDF generated successfully: {output_pdf_path}")
                 messagebox.showinfo("Success", f"PDF generated and saved to:\\n{output_pdf_path}")
            else:
                 self.status_label.config(text="Error: No pages generated.")
                 messagebox.showerror("Error", "PDF generation failed: No pages were created in the output document.")

        # ... (rest of exception handling and finally block remains mostly the same) ...
        except FileNotFoundError as e:
            self.status_label.config(text="Error: File not found.")
            messagebox.showerror("Error", str(e))
        except fitz.FileDataError as e: # Catch PyMuPDF specific errors reading/writing
            self.status_label.config(text="Error: PDF processing error.")
            messagebox.showerror("PDF Error", f"Error processing PDF files: {e}")
        except ImportError as e: # Should catch missing libs like reportlab, matplotlib
             self.status_label.config(text="Error: Missing libraries.")
             messagebox.showerror("Error", f"Required library not installed: {e}.")
        except Exception as e: # Catch any other unexpected error during PDF generation setup/finalization
            self.status_label.config(text="Error during PDF generation.")
            tb_str = traceback.format_exc()
            err_msg = str(e).lower()
            messagebox.showerror("Error", f"An unexpected error occurred:\\n{e}\\n\\nTraceback:\\n{tb_str}")
        finally:
             # Robust cleanup
             # plt.close('all') # No longer needed
             if source_doc: source_doc.close()
             if output_doc: output_doc.close()
             if 'table_pdf' in locals() and table_pdf: table_pdf.close(); table_pdf = None
             if 'temp_buffer' in locals() and temp_buffer and not temp_buffer.closed: temp_buffer.close()


# --- Run the Application ---
if __name__ == "__main__":
    root = tk.Tk()
    app = PdfGeneratorApp(root)
    root.mainloop() 