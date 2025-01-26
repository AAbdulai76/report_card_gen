import pandas as pd
from fpdf import FPDF
import os
import tkinter as tk
from tkinter import filedialog, messagebox


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

        # Dynamically detect subject columns
        subject_columns = [
            col for col in df.columns if col.strip().lower() != name_col.strip().lower()
        ]

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
            scores = [row[subject] for subject in subject_columns]
            scores = [
                float(score) if isinstance(score, (int, float)) else 0
                for score in scores
            ]

            class_scores = [round(score * 0.3) for score in scores]
            exam_scores = [round(score * 0.7) for score in scores]
            total_scores = scores
            student_positions = [positions[subject][idx] for subject in subject_columns]

            remarks = []
            for score in scores:
                if 80 <= score <= 100:
                    remarks.append("Excellent!")
                elif 70 <= score <= 79:
                    remarks.append("Very Good")
                elif 60 <= score <= 69:
                    remarks.append("Good")
                elif 50 <= score <= 59:
                    remarks.append("Developing")
                else:
                    remarks.append("Beginning")

            # Create PDF
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Times", style="B", size=30)
            pdf.cell(170, 10, txt="PROGRESS ACADEMY", border=1, ln=True, align="C")
            pdf.set_font("Courier", style="BU", size=20)
            pdf.cell(170, 10, txt="TERMINAL REPORT", ln=True, align="C")

            # Student Details
            pdf.ln(5)
            pdf.set_font("Times", size=12)
            pdf.cell(
                0,
                10,
                txt=f"NAME: {student_name}        ACADEMIC YEAR:__{year}__        CLASS: __{class_}__",
                ln=True,
            )
            pdf.cell(
                0,
                10,
                txt=f"NUMBER ON ROLL: __{number_on_roll}__       TERM: __{term}__       VACATION DATE: __{vacation_date}__",
                ln=True,
            )
            pdf.cell(
                0,
                10,
                txt=f"NEXT TERM BEGINS: __{next_term_begins}__       POSITION IN CLASS: __{row['Position']}__",
                ln=True,
            )

            # Table Headers
            pdf.set_font(family="Arial", style="B", size=10)
            pdf.set_text_color(0, 0, 128)  # Dark blue color (RGB)

            # Starting X position
            x_start = pdf.get_x()
            y_start = pdf.get_y()
            line_height = 6

            # First cell
            pdf.multi_cell(w=60, h=18, txt="SUBJECTS", border=1, align="C")
            pdf.set_xy(x_start + 60, y_start)

            # Second cell
            pdf.multi_cell(
                w=20, h=line_height, txt="CLASS\nSCORE\n(30%)", border=1, align="C"
            )
            pdf.set_xy(x_start + 80, y_start)

            # Third cell
            pdf.multi_cell(
                w=20, h=line_height, txt="EXAM\nSCORE\n(70%)", border=1, align="C"
            )
            pdf.set_xy(x_start + 100, y_start)

            # Fourth cell
            pdf.multi_cell(
                w=20, h=line_height, txt="TOTAL\nSCORE\n(100%)", border=1, align="C"
            )
            pdf.set_xy(x_start + 120, y_start)

            # Fifth cell
            pdf.multi_cell(
                w=20, h=line_height, txt="SUBJ.\nPOS\nCLASS", border=1, align="C"
            )
            pdf.set_xy(x_start + 140, y_start)

            # Last cell
            pdf.multi_cell(w=30, h=18, txt="REMARKS", border=1, align="C")

            # Move to next line
            pdf.set_xy(
                x_start, y_start + (line_height * 3)
            )  # Adjust this value based on the total height needed

            pdf.set_font("Arial", size=10)
            for i, subject in enumerate(subject_columns):
                pdf.cell(60, 10, txt=subject.title().upper(), border=1)
                pdf.cell(20, 10, txt=str(class_scores[i]), border=1, align="C")
                pdf.cell(20, 10, txt=str(exam_scores[i]), border=1, align="C")
                pdf.cell(20, 10, txt=str(total_scores[i]), border=1, align="C")
                pdf.cell(20, 10, txt=str(student_positions[i]), border=1, align="C")
                pdf.cell(30, 10, txt=remarks[i], border=1, ln=True, align="C")

            # Footer
            pdf.ln(10)
            pdf.set_font("Times", size=12)
            pdf.cell(
                0,
                10,
                txt="ATTENDANCE: ____________ OUT OF: ____________ PROMOTED TO/REPEATED IN: ___N/A___",
                ln=True,
            )
            pdf.cell(
                0,
                10,
                txt="CONDUCT: _________________________________________________________________________________",
                ln=True,
            )
            pdf.cell(
                0,
                10,
                txt="CLASS TEACHER'S REMARKS: _________________________________________________________________",
                ln=True,
            )
            pdf.cell(
                0,
                10,
                txt="HEAD TEACHER'S SIGNATURE: ________________________________________________________________",
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
            date_entry.get(),
            year_entry.get(),
            vacation_entry.get(),
            number_entry.get(),
            next_term_entry.get(),
            term_entry.get(),
            output_dir,
        )

    # Tkinter window
    root = tk.Tk()
    root.title("Report Card Generator")

    tk.Label(root, text="Excel File:").grid(row=0, column=0, sticky="e")
    file_entry = tk.Entry(root, width=50)
    file_entry.grid(row=0, column=1)
    tk.Button(root, text="Browse", command=browse_file).grid(row=0, column=2)

    labels = [
        "Date:",
        "Year:",
        "Vacation Date:",
        "Number on Roll:",
        "Next Term Begins:",
        "Term:",
    ]
    entries = []

    for i, label in enumerate(labels):
        tk.Label(root, text=label).grid(row=i + 1, column=0, sticky="e")
        entry = tk.Entry(root)
        entry.grid(row=i + 1, column=1)
        entries.append(entry)

    (
        date_entry,
        year_entry,
        vacation_entry,
        number_entry,
        next_term_entry,
        term_entry,
    ) = entries

    tk.Button(root, text="Generate Report Cards", command=generate_reports).grid(
        row=len(labels) + 1, column=0, columnspan=3
    )

    root.mainloop()


if __name__ == "__main__":
    create_gui()
