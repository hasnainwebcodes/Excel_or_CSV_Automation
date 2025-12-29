from flask import Flask, render_template, redirect, request, send_file
import pandas as pd
import os
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors


app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
CLEAN_FOLDER = "cleaned"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(CLEAN_FOLDER, exist_ok=True)

def dataframe_to_pdf(df, output_path, title_text="Cleaned Data Report"):
    page_width, page_height = A4

    pdf = SimpleDocTemplate(
        output_path,
        pagesize=A4,
        rightMargin=20,
        leftMargin=20,
        topMargin=30,
        bottomMargin=30
    )

    elements = []
    styles = getSampleStyleSheet()

    # 🔹 TITLE
    title_style = ParagraphStyle(
        "TitleStyle",
        parent=styles["Title"],
        alignment=1,  # center
        fontSize=18,
        spaceAfter=20
    )

    title = Paragraph(title_text, title_style)
    elements.append(title)
    elements.append(Spacer(1, 10))

    # 🔹 TABLE DATA
    data = [df.columns.tolist()] + df.astype(str).values.tolist()

    # 🔹 AUTO COLUMN WIDTH (FULL PAGE)
    available_width = page_width - pdf.leftMargin - pdf.rightMargin
    col_count = len(df.columns)
    col_widths = [available_width / col_count] * col_count

    table = Table(data, colWidths=col_widths, repeatRows=1)

    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),

        ("ALIGN", (0, 0), (-1, -1), "LEFT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),

        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),

        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 9),

        ("BOTTOMPADDING", (0, 0), (-1, 0), 10),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
    ]))

    elements.append(table)
    pdf.build(elements)

def read_csv_safely(file):
    try:
        file.seek(0)
        return pd.read_csv(
            file,
            engine="python",
            sep=None,
            on_bad_lines="skip"
        )
    except Exception as e:
        raise ValueError(f"CSV read error: {str(e)}")

@app.route("/", methods= ['GET','POST'])
def index():
    if request.method == "POST":
        file = request.files["file"]
        output_format= request.form.get("output_format", "csv")
        if not file or file.filename == "":
            return "No file uploaded", 400
        filename = file.filename.lower()

        if filename.endswith(".csv"):
            df = read_csv_safely(file)

        elif filename.endswith(".xlsx"):
            df = pd.read_excel(file, engine="openpyxl")

        else:
            return "Unsupported file format", 400
        df.columns = df.columns.str.strip().str.lower()
        df = df.drop_duplicates(ignore_index=True)
        df = df.dropna(how="all")

        for col in df.columns:
                df[col] = df[col].fillna("Unknown")
        if output_format == "xlsx":
            output_path = os.path.join(CLEAN_FOLDER, "cleaned_data.xlsx")
            df.to_excel(output_path, index=False)
        else:
            output_path = os.path.join(CLEAN_FOLDER, "cleaned_data.csv")
            df.to_csv(output_path, index=False)

        return send_file(output_path, as_attachment=True)
    return render_template("index.html")
@app.route("/merger/csv", methods= ['GET','POST'])
def csv_merger():
    if request.method == "POST":
        file1 = request.files["file1"]
        file2 = request.files["file2"]
        if not file1 or not file2 or file1.filename == "" or file2.filename == "":
            return "No file uploaded", 400
        filename1 = file1.filename.lower()
        filename2 = file2.filename.lower()
        try:
            if filename1.endswith(".csv"):
                df1 = read_csv_safely(file1)
            if filename2.endswith(".csv"):
                df2 = read_csv_safely(file2)
        except Exception as e :
            return "Can't open one or both CSV files", 400
        df1.columns = df1.columns.str.strip().str.lower()
        df2.columns = df2.columns.str.strip().str.lower()
        df3 = pd.concat([df1, df2], ignore_index=True)
        df3 = df3.drop_duplicates(ignore_index=True)

        output_path = os.path.join(CLEAN_FOLDER, "merged_data.csv")
        df3.to_csv(output_path, index=False)

        return send_file(output_path, as_attachment=True)

    return render_template("csvmerger.html")
@app.route("/merger/excel", methods=['GET', 'POST'])
def excel_merger():
    if request.method == "POST":
        file1 = request.files["file1"]
        file2 = request.files["file2"]
        if not file1 or not file2 or file1.filename == "" or file2.filename == "":
            return "No file uploaded", 400
        filename1 = file1.filename.lower()
        filename2 = file2.filename.lower()
        try:
            df1 = pd.read_excel(file1, engine="openpyxl")
            df2 = pd.read_excel(file2, engine="openpyxl")
        except Exception as e :
            return "Can't open one or both files" , 500 
        df1.columns = df1.columns.str.strip().str.lower()
        df2.columns = df2.columns.str.strip().str.lower()
        df3 = pd.concat([df1, df2], ignore_index=True)
        df3 = df3.drop_duplicates(ignore_index=True)

        output_path = os.path.join(CLEAN_FOLDER, "merged_data..xlsx")
        df3.to_excel(output_path, index=False)

        return send_file(output_path, as_attachment=True)
        return "Hi"
    return render_template("excelmerger.html")
@app.route("/converter/pdf", methods= ['GET', 'POST'])
def excel_to_pdf():
    if request.method == "POST":
        file = request.files.get("file")

        if not file or file.filename == "":
            return "No file uploaded", 400

        filename = file.filename.lower()

        if filename.endswith(".csv"):
            df = read_csv_safely(file)

        elif filename.endswith(".xlsx"):
            df = pd.read_excel(file, engine="openpyxl")

        else:
            return "Only CSV or Excel allowed", 400

        # 🔹 CLEANING
        df.columns = df.columns.str.strip().str.lower()
        df = df.drop_duplicates(ignore_index=True)
        df = df.dropna(how="all")
        df = df.fillna("Unknown")

        output_path = os.path.join(CLEAN_FOLDER, "cleaned_data.pdf")
        dataframe_to_pdf(df, output_path)

        return send_file(output_path, as_attachment=True)
    return render_template("exceltopdf.html")

app.run(port=3010)