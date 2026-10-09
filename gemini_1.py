import os
import re
import pandas as pd
import pdfplumber

# Folder containing your USTER TEST Reports
FOLDER_PATH = r"C:\Users\info\Dropbox\Thesis\PhD\USTER TEST Reports\USTER TEST NEW"
OUTPUT_EXCEL = r"C:\Users\info\Dropbox\Thesis\PhD\USTER TEST Reports\USTER_Combined_Summary.xlsx"

# Regexes for the metadata header line, e.g.:
#   "Stil 2026 Örnek ID 26005 Nom. numara Nec 12.5 Nom. büküm 0 T/m"
#   "Stil 2026 Örnek ID 26047 Nom. numara Nec 10 Nom. büküm 0 T/m"
# NOTE: in the actual extracted PDF text, "Örnek ID" and "Nom. numara" sit on
# the SAME line as their values (not on the next line), so the old regexes
# (which looked for a line break right after "ID" / "numara") never matched
# and always returned blank metadata. Fixed below.
ORNEK_ID_RE = re.compile(r"Örnek\s+ID\s+(\S+)")
NOM_NUM_RE = re.compile(r"Nom\.\s+numara\s+(.+?)\s+Nom\.\s+büküm")


def parse_uster_pdf(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        full_text = "\n".join([page.extract_text() or "" for page in pdf.pages])

    lines = full_text.splitlines()
    rows = []

    # Track the "current" Örnek ID / Nom. numara as we scan down the file.
    # This means a single PDF that contains several samples back-to-back
    # (e.g. one block for Örnek ID 26005, later a block for 26047 with
    # Nom. numara Nec 10) will correctly tag each data row with the header
    # that precedes it, instead of only ever using the first header found
    # in the whole document.
    current_ornek_id = ""
    current_nom_num = ""

    for line in lines:
        line_clean = line.strip()

        oid_match = ORNEK_ID_RE.search(line_clean)
        if oid_match:
            current_ornek_id = oid_match.group(1)

        nn_match = NOM_NUM_RE.search(line_clean)
        if nn_match:
            current_nom_num = nn_match.group(1).strip()

        # Look for numerical data rows (starting with row number 1, 2, etc.)
        tokens = line_clean.split()
        if len(tokens) >= 11 and tokens[0].isdigit():
            # Check if second item looks like a numeric metric (e.g. U%)
            try:
                float(tokens[1])
            except ValueError:
                continue

            row_data = {
                "File Name": os.path.basename(pdf_path),
                "Örnek ID": current_ornek_id,
                "Nom. numara": current_nom_num,
                "Nr": tokens[0],
                "U%": tokens[1],
                "CVm%": tokens[2],
            }

            # If 12 numeric tokens, Index is included; if 11, Index is missing
            if len(tokens) >= 12:
                row_data["Index"] = tokens[3]
                vals = tokens[4:]
            else:
                row_data["Index"] = ""
                vals = tokens[3:]

            if len(vals) >= 8:
                row_data["Ince -40%/km"] = vals[0]
                row_data["Ince -50%/km"] = vals[1]
                row_data["Kalin +35%/km"] = vals[2]
                row_data["Kalin +50%/km"] = vals[3]
                row_data["Neps +200%/km"] = vals[4]
                row_data["Neps +280%/km"] = vals[5]
                row_data["Rel. Num. ± %"] = vals[6]
                row_data["H"] = vals[7]
                row_data["sh"] = vals[8] if len(vals) > 8 else ""

                rows.append(row_data)

    return rows


def batch_process():
    all_records = []

    pdf_files = [f for f in os.listdir(FOLDER_PATH) if f.lower().endswith('.pdf')]
    print(f"Found {len(pdf_files)} PDF files to process...")

    for file_name in pdf_files:
        pdf_path = os.path.join(FOLDER_PATH, file_name)
        try:
            records = parse_uster_pdf(pdf_path)
            all_records.extend(records)
            print(f"Processed: {file_name} ({len(records)} row(s))")
        except Exception as e:
            print(f"Error reading {file_name}: {e}")

    # Export to Excel
    if all_records:
        df = pd.DataFrame(all_records)
        df.to_excel(OUTPUT_EXCEL, index=False)
        print(f"\nDone! Extracted {len(all_records)} total rows into:\n{OUTPUT_EXCEL}")
    else:
        print("\nNo records were extracted. Please check the PDF files.")


if __name__ == "__main__":
    batch_process()
