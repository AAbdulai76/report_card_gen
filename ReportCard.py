import pandas as pd
from fpdf import FPDF
import os
import tkinter as tk
from tkinter import filedialog, messagebox


def ordinal(n):
    """Convert an integer into its ordinal representation."""
    if 11 <= n % 100 <= 13:
        suffix = "th"
    else:
        suffix = {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th")
    return f"{n}{suffix}"


# Function to generate report cards
def generate_report_cards_with_positions(
    excel_file,
    class_,
    year,
    vacation_date,
    number_on_roll,
    next_term_begins,
    term,
    output_dir,
):
    try:
        # Create the output directory if it doesn't exist
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        # Load the Excel file
        df = pd.read_excel(excel_file)

        # Replace missing (NaN) values with 0
        df = df.fillna(0)

        # Ensure the 'Name' column exists
        name_col = next(
            (col for col in df.columns if col.strip().lower() == "name"), None
        )
        if not name_col:
            raise ValueError("Excel file must have a 'Name' column for student names.")

        # Ensure the required columns exist
        required_columns = ["Conduct", "Remarks"]
        for col in required_columns:
            if col not in df.columns:
                raise ValueError(f"Excel file must have a '{col}' column.")

        # Dynamically detect subject columns (exclude non-numeric columns like 'Conduct' and 'Remarks')
        subject_columns = [
            col
            for col in df.columns
            if col.strip().lower()
            not in [name_col.strip().lower(), "conduct", "remarks"]
        ]

        # Ensure subject columns contain only numeric data
        for col in subject_columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

        # Calculate total scores for each student
        df["Total_Score"] = df[subject_columns].sum(axis=1)

        # Calculate overall positions
        df["Position"] = (
            df["Total_Score"].rank(ascending=False, method="min").astype(int)
        )

        # Sort dataframe based on positions (optional)
        df = df.sort_values(by="Position")

        # Calculate positions for each subject
        positions = {}
        for subject in subject_columns:
            positions[subject] = (
                df[subject].rank(ascending=False, method="min").astype(int)
            )

        # Loop through each student and generate a report card
        for idx, row in df.iterrows():
            student_name = row[name_col]
            conduct = row["Conduct"]  # Read the Conduct column
            remarks = row["Remarks"]  # Read the Remarks column
            scores = [row[subject] for subject in subject_columns]
            scores = [
                float(score) if isinstance(score, (int, float)) else 0
                for score in scores
            ]

            class_scores = [round(score * 0.3) for score in scores]
            exam_scores = [round(score * 0.7) for score in scores]
            total_scores = scores
            student_positions = [positions[subject][idx] for subject in subject_columns]

            remarks_list = []
            for score in scores:
                if 80 <= score <= 100:
                    remarks_list.append("Excellent!")
                elif 70 <= score <= 79:
                    remarks_list.append("Very Good")
                elif 60 <= score <= 69:
                    remarks_list.append("Good")
                elif 50 <= score <= 59:
                    remarks_list.append("Developing")
                else:
                    remarks_list.append("Beginning")

            # Create PDF
            pdf = FPDF()
            pdf.set_left_margin(10)
            pdf.set_right_margin(10)
            pdf.add_page()

            # Add a decorative rectangle around the report card
            pdf.set_draw_color(0, 0, 0)  # Black border color
            pdf.set_line_width(1)  # Outer border thickness
            pdf.rect(
                x=8, y=8, w=194, h=281
            )  # Outer rectangle dimensions (A4 page with 8mm margins)

            # Add an inner rectangle for a layered effect
            pdf.set_draw_color(0, 0, 0)  # Black border color
            pdf.set_line_width(0.5)  # Inner border thickness
            pdf.rect(
                x=10, y=10, w=190, h=277
            )  # Inner rectangle dimensions (A4 page with 10mm margins

            # Register the Rockwell fonts
            pdf.add_font(
                "Rockwell", style="", fname="./fonts/ROCK.TTF", uni=True
            )  # Regular
            pdf.add_font(
                "Rockwell", style="B", fname="./fonts/ROCKB.TTF", uni=True
            )  # Bold
            pdf.add_font(
                "Rockwell", style="I", fname="./fonts/ROCKI.TTF", uni=True
            )  # Italic
            pdf.add_font(
                "Rockwell", style="BI", fname="./fonts/ROCKBI.TTF", uni=True
            )  # Bold Italic

            # Add the school's logo on the left
            logo_path = "./logo.png"
            pdf.image(
                logo_path, x=12, y=13, w=35, h=35
            )  # Adjusted position (inside the border)

            # Add the school's logo on the right
            pdf.image(
                logo_path, x=162, y=13, w=35, h=35
            )  # Adjusted position (inside the border)

            # Add the school name using Rockwell Bold font
            pdf.set_font("Rockwell", style="B", size=32)  # Use bold Rockwell font
            pdf.cell(0, 10, txt="PROGRESS ACADEMY", ln=True, align="C")

            # Add the school address and contact details
            pdf.set_font("Rockwell", style="", size=12)  # Use regular Rockwell font
            pdf.cell(0, 10, txt="P. O. BOX TL 481, TAMALE - GHANA", ln=True, align="C")
            pdf.cell(0, 10, txt="TELEPHONE: +233 24 385 7907", ln=True, align="C")

            # Add the terminal report title without a background color
            pdf.set_text_color(0, 0, 0)  # Set text color to black
            pdf.set_font("Rockwell", style="B", size=15)
            pdf.cell(0, 10, txt="TERMINAL REPORT", ln=True, align="C")

            # Add the student's name in the heading
            pdf.ln(5)
            pdf.set_font("Times", style="B", size=17)
            pdf.cell(
                0, 10, txt=f"[ {student_name.upper()} ]", border=1, ln=True, align="C"
            )

            # Student Details
            pdf.ln(5)

            # Define column widths (adjusted to fit within 190mm)
            col_widths = [
                50,
                50,
                50,
            ]  # Adjusted column widths to ensure they fit within the page

            # Row 1: CLASS, ACADEMIC YEAR, TERM
            pdf.set_font("Rockwell", style="", size=12)  # Regular font for name tags
            pdf.cell(col_widths[0], 10, txt="CLASS:", border=0, align="L")
            pdf.set_font("Times", style="B", size=12)  # Bold font for values
            pdf.cell(col_widths[0], 10, txt=class_, border=0, align="L")

            pdf.set_font("Rockwell", style="", size=12)  # Regular font for name tags
            pdf.cell(col_widths[1], 10, txt="ACADEMIC YEAR:", border=0, align="L")
            pdf.set_font("Times", style="B", size=12)  # Bold font for values
            pdf.cell(col_widths[1], 10, txt=year, border=0, align="L")

            pdf.set_font("Rockwell", style="", size=12)  # Regular font for name tags
            pdf.cell(col_widths[2], 10, txt="TERM:", border=0, align="L")
            pdf.set_font("Times", style="B", size=12)  # Bold font for values
            pdf.cell(col_widths[2], 10, txt=term, border=0, ln=True, align="L")

            # Row 2: NUMBER ON ROLL, POSITION IN CLASS
            pdf.set_font("Rockwell", style="", size=12)  # Regular font for name tags
            pdf.cell(col_widths[0], 10, txt="NUMBER ON ROLL:", border=0, align="L")
            pdf.set_font("Times", style="B", size=12)  # Bold font for values
            pdf.cell(col_widths[0], 10, txt=str(number_on_roll), border=0, align="L")

            pdf.set_font("Rockwell", style="", size=12)  # Regular font for name tags
            pdf.cell(col_widths[1], 10, txt="POSITION IN CLASS:", border=0, align="L")
            pdf.set_font("Times", style="B", size=12)  # Bold font for values
            pdf.cell(
                col_widths[1],
                10,
                txt=ordinal(row["Position"]),
                border=0,
                ln=True,
                align="L",
            )

            # Row 3: NEXT TERM RE-OPENS, VACATION DATE
            pdf.set_font("Rockwell", style="", size=12)  # Regular font for name tags
            pdf.cell(col_widths[0], 10, txt="NEXT TERM RE-OPENS:", border=0, align="L")
            pdf.set_font("Times", style="B", size=12)  # Bold font for values
            pdf.cell(col_widths[0], 10, txt=next_term_begins, border=0, align="L")

            pdf.set_font("Rockwell", style="", size=12)  # Regular font for name tags
            pdf.cell(col_widths[0], 10, txt="VACATION DATE:", border=0, align="L")
            pdf.set_font("Times", style="B", size=12)  # Bold font for values
            pdf.cell(col_widths[0], 10, txt=vacation_date, border=0, ln=True, align="L")

            # Table Headers
            pdf.set_font(family="Arial", style="B", size=10)
            pdf.set_text_color(0, 0, 0)  # Black color (RGB)

            # Starting X position
            x_start = pdf.get_x()
            y_start = pdf.get_y()
            line_height = 6

            # First cell (SUBJECTS)
            pdf.multi_cell(
                w=70, h=18, txt="SUBJECTS", border=1, align="C"
            )  # Increased width
            pdf.set_xy(x_start + 70, y_start)

            # Second cell (CLASS SCORE)
            pdf.multi_cell(
                w=18, h=line_height, txt="CLASS\nSCORE\n(30%)", border=1, align="C"
            )
            pdf.set_xy(x_start + 88, y_start)

            # Third cell (EXAM SCORE)
            pdf.multi_cell(
                w=18, h=line_height, txt="EXAM\nSCORE\n(70%)", border=1, align="C"
            )
            pdf.set_xy(x_start + 106, y_start)

            # Fourth cell (TOTAL SCORE)
            pdf.multi_cell(
                w=18, h=line_height, txt="TOTAL\nSCORE\n(100%)", border=1, align="C"
            )
            pdf.set_xy(x_start + 124, y_start)

            # Fifth cell (SUBJ. POS CLASS)
            pdf.multi_cell(
                w=18, h=line_height, txt="SUBJ.\nPOS.\nCLASS", border=1, align="C"
            )
            pdf.set_xy(x_start + 142, y_start)

            # Last cell (REMARKS)
            pdf.multi_cell(
                w=48, h=18, txt="REMARKS", border=1, align="C"
            )  # Increased width

            # Move to next line
            pdf.set_xy(x_start, y_start + (line_height * 3))

            # Table Rows
            pdf.set_font("Arial", size=10)
            for i, subject in enumerate(subject_columns):
                pdf.cell(70, 10, txt=subject.title().upper(), border=1)  # SUBJECTS
                pdf.cell(
                    18, 10, txt=str(class_scores[i]), border=1, align="C"
                )  # CLASS SCORE
                pdf.cell(
                    18, 10, txt=str(exam_scores[i]), border=1, align="C"
                )  # EXAM SCORE
                pdf.cell(
                    18, 10, txt=str(total_scores[i]), border=1, align="C"
                )  # TOTAL SCORE
                pdf.cell(
                    18, 10, txt=ordinal(student_positions[i]), border=1, align="C"
                )  # SUBJ. POS CLASS with ordinal
                pdf.cell(
                    48, 10, txt=remarks_list[i], border=1, ln=True, align="C"
                )  # REMARKS

            # Footer
            pdf.ln(10)
            pdf.set_font("Rockwell", style="", size=12)
            pdf.cell(
                0,
                10,
                txt="ATTENDANCE: ____________ OUT OF: ____________ PROMOTED TO/REPEATED IN:  ___N/A___ ",
                ln=True,
            )
            pdf.cell(
                0,
                10,
                txt=f"CONDUCT:  {conduct}",
                ln=True,
            )
            pdf.cell(
                0,
                10,
                txt=f"CLASS TEACHER'S REMARKS:    {remarks}",
                ln=True,
            )
            pdf.cell(
                0,
                10,
                txt="CLASS TEACHER'S SIGNATURE: ______________ HEAD TEACHER'S SIGNATURE: ______________",
                ln=True,
            )
            filename = os.path.join(
                output_dir, f"{student_name.replace(' ', '_')}_report_card.pdf"
            )
            pdf.output(filename)

        messagebox.showinfo("Success", "Report cards generated successfully!")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")


# GUI
def create_gui():
    # Tkinter window
    root = tk.Tk()
    root.title("Report Card Generator")

    def browse_file():
        filename = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        file_entry.delete(0, tk.END)
        file_entry.insert(0, filename)

    def generate_reports():
        file = file_entry.get()
        if not file:
            messagebox.showwarning("Input Error", "Please select an Excel file.")
            return
        output_dir = filedialog.askdirectory()
        if not output_dir:
            messagebox.showwarning("Input Error", "Please select an output directory.")
            return
        generate_report_cards_with_positions(
            file,
            class_entry.get(),
            year_entry.get(),
            vacation_entry.get(),
            number_entry.get(),
            next_term_entry.get(),
            term_entry.get(),
            output_dir,
        )

    # Labels and entries for the GUI
    labels = [
        "Year:",
        "Vacation Date:",
        "Number on Roll:",
        "Next Term Begins:",
        "Term:",
        "Class:",
    ]
    entries = []

    for i, label in enumerate(labels):
        tk.Label(root, text=label).grid(row=i + 1, column=0, sticky="e")
        entry = tk.Entry(root)
        entry.grid(row=i + 1, column=1)
        entries.append(entry)

    (
        year_entry,
        vacation_entry,
        number_entry,
        next_term_entry,
        term_entry,
        class_entry,
    ) = entries

    tk.Label(root, text="Excel File:").grid(row=0, column=0, sticky="e")
    file_entry = tk.Entry(root, width=50)
    file_entry.grid(row=0, column=1)
    tk.Button(root, text="Browse", command=browse_file).grid(row=0, column=2)

    tk.Button(root, text="Generate Report Cards", command=generate_reports).grid(
        row=len(labels) + 1, column=0, columnspan=3
    )

    root.mainloop()


if __name__ == "__main__":
    create_gui()
