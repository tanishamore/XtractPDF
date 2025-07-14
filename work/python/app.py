from flask import Flask, session, render_template, request, send_file, redirect, url_for
import pdfplumber
import pandas as pd
import os

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-key")

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "outputs"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        file = request.files["pdf"]
        password = request.form.get("password", "").strip()

        # Check file extension (only PDF allowed)
        if not file.filename.endswith(".pdf"):
            return "<h3>Only PDF files are allowed.</h3><a href='/'>Try Again</a>"

        # Check file size limit (max 5 MB)
        file.seek(0, os.SEEK_END)
        size = file.tell()
        file.seek(0)

        if size > 5 * 1024 * 1024:
            return "<h3>File too large (max 5 MB allowed).</h3><a href='/'>Try Again</a>"

        if file and file.filename:
            print("✅ File received:", file.filename)
            filepath = os.path.join(UPLOAD_FOLDER, file.filename)
            file.save(filepath)
            print("✅ File saved to:", filepath)

            try:
                with pdfplumber.open(filepath, password=password) as pdf:
                    print("✅ PDF opened")
                    output_path = os.path.join(OUTPUT_FOLDER, "output.xlsx")

                    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
                        last_columns = None
                        sheet_counter = 1
                        buffer_df = pd.DataFrame()
                        preview_df = None

                        for page_num, page in enumerate(pdf.pages):
                            print(f"⏳ Processing page {page_num + 1}")
                            table = page.extract_table()
                            if table:
                                df = pd.DataFrame(table[1:], columns=table[0])

                                if preview_df is None and not df.empty:
                                    preview_df = df.head(10)
                                    preview_df = preview_df.applymap(lambda x: str(x).replace('\n', '<br>') if pd.notnull(x) else x)

                                if last_columns is not None and df.columns.tolist() == last_columns:
                                    buffer_df = pd.concat([buffer_df, df], ignore_index=True)
                                else:
                                    if not buffer_df.empty:
                                        sheet_name = f"Sheet{sheet_counter}"
                                        buffer_df.to_excel(writer, sheet_name=sheet_name, index=False)
                                        sheet_counter += 1

                                    buffer_df = df
                                    last_columns = df.columns.tolist()

                        if not buffer_df.empty:
                            sheet_name = f"Sheet{sheet_counter}"
                            buffer_df.to_excel(writer, sheet_name=sheet_name, index=False)

                if preview_df is not None:
                    session["preview_html"] = preview_df.to_html(classes="table table-striped", index = False, escape=False)
                    session["excel_path"] = output_path
                    print("✅ Preview HTML stored in session")

                    return redirect(url_for("preview"))
                
                return "<h3>No tabular data found in the PDF.</h3><a href='/'>Try Again</a>"

            except Exception as e:
                print("❌ Error:", e)
                return f"<h3>Error: {str(e)}</h3><a href='/'>Try Again</a>"

        return "<h3>No file uploaded.</h3><a href='/'>Try Again</a>"

    return render_template("index.html")

@app.route("/preview")
def preview():
    preview_html = session.get("preview_html","<p>No preview available.</p>")
    return render_template("preview.html", preview_table = preview_html)

@app.route("/download")
def download():
    excel_path = session.get("excel_path")
    if excel_path and os.path.exists(excel_path):
        return send_file(excel_path, as_attachment=True)
    return "<h3>Excel file not found.</h3><a href='/'>Try Again</a>"

if __name__ == "__main__":
    app.run(debug=True)