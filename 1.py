import tkinter as tk
from tkinter import filedialog, messagebox, Text, Scale
from PIL import Image, ImageDraw, ImageFont, ImageTk
import pandas as pd
import os
import random
import re

# --- Global variables ---
template_path = ""
excel_path = ""
output_folder = "generated_certificates"
FONT_REGULAR_PATH = "arial.ttf"
FONT_BOLD_PATH = "arialbd.ttf"
FONT_SIZE = 50

# CertificateGeneratorApp class me koi badlav nahi, woh waisi hi rahegi.
class CertificateGeneratorApp:
    # ... (code from previous answer, no changes here)
    def __init__(self, root):
        self.root = root
        self.root.title("Chess Certificate Generator")
        self.root.geometry("700x550")

        frame = tk.Frame(root, padx=15, pady=15)
        frame.pack(padx=10, pady=10, fill="both", expand=True)

        self.btn_template = tk.Button(frame, text="1. Upload Certificate Template (Image)", command=self.upload_template)
        self.btn_template.pack(fill="x", pady=5)
        self.lbl_template = tk.Label(frame, text="No template selected", fg="red")
        self.lbl_template.pack()

        self.btn_excel = tk.Button(frame, text="2. Upload Excel Data (.xlsx)", command=self.upload_excel)
        self.btn_excel.pack(fill="x", pady=5)
        self.lbl_excel = tk.Label(frame, text="No Excel file selected", fg="red")
        self.lbl_excel.pack()

        tk.Label(frame, text="3. Certificate Text (Use **text** for bold)", font=("Arial", 10, "bold")).pack(pady=(15, 5), anchor="w")
        
        self.txt_input = Text(frame, height=10, wrap="word", font=("Arial", 12))
        self.txt_input.pack(fill="both", expand=True, pady=5)
        
        # Default text updated with **Mr.** example
        default_text = (
            "This is to certify that {Name} S/o **Mr.** {Father_Name} "
            "has secured {Position} position in U-17 Girls category.\n\n"
            "He has obtained {Points} points in {Rounds} rounds in **State Chess Championship** "
            "held on 28th and 29th June, 2025 at KR Mangalam World School, "
            "Sector 2, Bahadurgarh, Jhajjar, Haryana."
        )
        self.txt_input.insert(tk.END, default_text)

        self.btn_preview = tk.Button(frame, text="Generate Preview & Adjust", command=self.generate_preview, state="disabled")
        self.btn_preview.pack(pady=20, fill="x", ipady=10)

    def upload_template(self):
        global template_path
        template_path = filedialog.askopenfilename(filetypes=[("Image Files", "*.png *.jpg *.jpeg")])
        if template_path:
            self.lbl_template.config(text=os.path.basename(template_path), fg="green")
            self.check_inputs()

    def upload_excel(self):
        global excel_path
        excel_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if excel_path:
            self.lbl_excel.config(text=os.path.basename(excel_path), fg="green")
            self.check_inputs()

    def check_inputs(self):
        if template_path and excel_path:
            self.btn_preview.config(state="normal")

    def generate_preview(self):
        user_text = self.txt_input.get("1.0", tk.END).strip()
        if not user_text:
            messagebox.showerror("Error", "Please enter the text for the certificate.")
            return

        try:
            df = pd.read_excel(excel_path)
            required_cols = ['Name', 'Father_Name', 'Position', 'Points', 'Rounds']
            missing_cols = [col for col in required_cols if col not in df.columns]
            if missing_cols:
                messagebox.showerror("Excel Error", f"Missing columns in Excel: {', '.join(missing_cols)}")
                return
                
            preview_data = df.sample(1).iloc[0]
            PreviewWindow(self.root, preview_data, user_text, df)

        except FileNotFoundError as e:
            messagebox.showerror("Font Error", f"Font file not found. Make sure '{FONT_REGULAR_PATH}' and '{FONT_BOLD_PATH}' are in the same folder.\nDetails: {e}")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")


class PreviewWindow:
    # --- YEH CLASS ABHI BHI WAHI HAI, SIRF draw_text_on_image FUNCTION ME BADLAV HAI ---
    def __init__(self, parent, preview_data, user_text, full_dataframe):
        self.parent = parent
        self.preview_data = preview_data
        self.user_text = user_text
        self.full_df = full_dataframe
        self.base_image = Image.open(template_path).convert("RGB")
        
        self.text_x = self.base_image.width * 0.1
        self.text_y = self.base_image.height / 2
        self.font_size = FONT_SIZE
        self.line_spacing = 1.6
        self.max_text_width = self.base_image.width * 0.8

        self.window = tk.Toplevel(self.parent)
        self.window.title("Adjust Text Position & Width")
        self.window.geometry("1200x700")
        
        control_frame = tk.Frame(self.window, padx=15, pady=15, bg="#f0f0f0")
        control_frame.pack(side=tk.RIGHT, fill=tk.Y)
        canvas_frame = tk.Frame(self.window)
        canvas_frame.pack(side=tk.LEFT, expand=True, fill=tk.BOTH)

        self.canvas = tk.Canvas(canvas_frame, scrollregion=(0, 0, self.base_image.width, self.base_image.height))
        hbar = tk.Scrollbar(canvas_frame, orient=tk.HORIZONTAL, command=self.canvas.xview)
        hbar.pack(side=tk.BOTTOM, fill=tk.X)
        vbar = tk.Scrollbar(canvas_frame, orient=tk.VERTICAL, command=self.canvas.yview)
        vbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.canvas.config(xscrollcommand=hbar.set, yscrollcommand=vbar.set)
        self.canvas.pack(expand=True, fill=tk.BOTH)

        tk.Label(control_frame, text="Set Text Width", font=("Arial", 14, "bold"), bg="#f0f0f0").pack(pady=(20, 5))
        self.width_slider = Scale(control_frame, from_=100, to=self.base_image.width, orient=tk.HORIZONTAL,
                                  command=self.on_width_change, length=200)
        self.width_slider.set(self.max_text_width)
        self.width_slider.pack(fill="x", pady=5)

        tk.Label(control_frame, text="Adjust Position", font=("Arial", 14, "bold"), bg="#f0f0f0").pack(pady=10)
        tk.Button(control_frame, text="⬆ Up", command=lambda: self.move_text(0, -10)).pack(fill="x", pady=3)
        tk.Button(control_frame, text="⬇ Down", command=lambda: self.move_text(0, 10)).pack(fill="x", pady=3)
        tk.Button(control_frame, text="⬅ Left", command=lambda: self.move_text(-10, 0)).pack(fill="x", pady=3)
        tk.Button(control_frame, text="➡ Right", command=lambda: self.move_text(10, 0)).pack(fill="x", pady=3)

        tk.Label(control_frame, text="Adjust Size", font=("Arial", 14, "bold"), bg="#f0f0f0").pack(pady=(20, 10))
        tk.Button(control_frame, text="➕ Increase Font", command=lambda: self.adjust_font(2)).pack(fill="x", pady=3)
        tk.Button(control_frame, text="➖ Decrease Font", command=lambda: self.adjust_font(-2)).pack(fill="x", pady=3)
        
        self.btn_generate_all = tk.Button(control_frame, text="✅ Done & Generate All", bg="#28a745", fg="white", font=("Arial", 12, "bold"), command=self.generate_all_certificates)
        self.btn_generate_all.pack(side=tk.BOTTOM, fill="x", ipady=10, pady=20)
        
        self.update_preview()
    
    # --- YEH FUNCTION POORA UPDATE HO GAYA HAI ---
    def draw_text_on_image(self, image, data):
        draw = ImageDraw.Draw(image)
        try:
            font_regular = ImageFont.truetype(FONT_REGULAR_PATH, self.font_size)
            font_bold = ImageFont.truetype(FONT_BOLD_PATH, self.font_size)
        except IOError as e:
            messagebox.showerror("Font Error", f"Could not load fonts. Details: {e}")
            return image

        current_x = self.text_x
        current_y = self.text_y
        space_width = font_regular.getlength(" ")

        # Regex to split text by placeholders OR by **bolded text**
        # This will give us a list of normal text, placeholders, and bolded text
        splitter = re.compile(r'(\{[^}]+\}|\*\*.*?\*\*)')

        for line in self.user_text.split('\n'):
            if line == '':
                current_y += self.font_size * self.line_spacing
                current_x = self.text_x
                continue
            
            parts = [p for p in splitter.split(line) if p]

            for part in parts:
                font = font_regular
                text_to_draw = part
                
                # Check what kind of part it is
                if part.startswith('{') and part.endswith('}'):
                    # It's a placeholder
                    col_name = part.strip('{}')
                    text_to_draw = str(data.get(col_name, ''))
                    font = font_bold
                elif part.startswith('**') and part.endswith('**'):
                    # It's user-defined bold text
                    text_to_draw = part.strip('*')
                    font = font_bold
                
                # Now, wrap the text word by word
                words = text_to_draw.split()
                for word in words:
                    word_width = font.getlength(word)
                    
                    if current_x + word_width > self.text_x + self.max_text_width:
                        current_y += self.font_size * self.line_spacing
                        current_x = self.text_x
                    
                    draw.text((current_x, current_y), word, font=font, fill="black")
                    current_x += word_width + space_width
            
            # After each full line from the text box, go to the next line on the image
            current_y += self.font_size * self.line_spacing
            current_x = self.text_x

        return image

    def on_width_change(self, value):
        self.max_text_width = int(value)
        self.update_preview()
        
    def update_preview(self):
        image_with_text = self.base_image.copy()
        image_with_text = self.draw_text_on_image(image_with_text, self.preview_data)
        self.tk_image = ImageTk.PhotoImage(image_with_text)
        self.canvas.create_image(0, 0, anchor="nw", image=self.tk_image)

    def move_text(self, dx, dy):
        self.text_x += dx
        self.text_y += dy
        self.update_preview()
    
    def adjust_font(self, d_size):
        self.font_size += d_size
        if self.font_size < 10: self.font_size = 10
        self.update_preview()

    def generate_all_certificates(self):
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)
        
        try:
            total = len(self.full_df)
            for index, row in self.full_df.iterrows():
                cert_image = self.base_image.copy()
                cert_image = self.draw_text_on_image(cert_image, row)
                student_name = str(row.get('Name', f'user_{index}')).replace(" ", "_")
                output_path = os.path.join(output_folder, f"certificate_{student_name}.png")
                cert_image.save(output_path)
                print(f"({index+1}/{total}) Generated: {output_path}")

            messagebox.showinfo("Success!", f"{total} certificates generated in '{output_folder}' folder.")
            self.window.destroy()
        except Exception as e:
            messagebox.showerror("Generation Failed", f"An error occurred: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = CertificateGeneratorApp(root)
    root.mainloop()