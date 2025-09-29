import os
import re
import pdfplumber
import pandas as pd
import logging
from .config import COMPANY_KEYWORDS, DEBUG_SAVE_RAW

def normalize_filename(filename: str) -> str:
    """Normalize a filename for company detection."""
    normalized = filename.lower()
    normalized = normalized.replace("_", " ").replace("-", " ")
    normalized = normalized.replace("&", " & ")
    normalized = re.sub(r"\s+", " ", normalized)
    return normalized.strip()

def detect_company_from_filename(file_path: str) -> str:
    """Detect company name from a file path using keywords."""
    filename = os.path.basename(file_path)
    normalized = normalize_filename(filename)
    for keyword, company_name in COMPANY_KEYWORDS.items():
        if keyword.lower() in normalized:
            return company_name
    return "Unknown Company"

def extract_pdf_to_dataframe(pdf_path: str) -> pd.DataFrame:
    """Extracts text from a PDF and returns a DataFrame preserving layout."""
    all_data = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                words = page.extract_words()
                if not words:
                    continue
                words.sort(key=lambda w: (w["top"], w["x0"]))
                rows = []
                current_row_y = None
                current_row = []
                for word in words:
                    y = round(word["top"], 1)
                    if current_row_y is None:
                        current_row_y = y
                    if abs(y - current_row_y) > 3:
                        rows.append(current_row)
                        current_row = []
                        current_row_y = y
                    current_row.append((word["x0"], word["text"]))
                if current_row:
                    rows.append(current_row)
                all_x = sorted(set([round(x0) for row in rows for x0, _ in row]))
                page_data = []
                for row in rows:
                    row_data = [""] * len(all_x)
                    for x0, text in row:
                        idx = min(range(len(all_x)), key=lambda i: abs(all_x[i] - round(x0)))
                        if row_data[idx]:
                            row_data[idx] += " " + text
                        else:
                            row_data[idx] = text
                    page_data.append(row_data)
                all_data.extend(page_data)
    except Exception as e:
        logging.error(f"Error extracting PDF {pdf_path}: {e}")
        return pd.DataFrame()
    if not all_data:
        return pd.DataFrame()
    return pd.DataFrame(all_data)

def filter_sanco(df: pd.DataFrame) -> pd.DataFrame:
    """Filter columns for Sanco company."""
    try:
        selected = df.iloc[:, [18, 36, 41, 46, 49]]  # S, AK, AP, AU, AX
        selected.columns = ["Col_S", "Col_AK", "Col_AP", "Col_AU", "Col_AX"]
        return selected
    except Exception as e:
        logging.warning(f"Sanco filter failed: {e}")
        return pd.DataFrame()

COMPANY_FILTERS = {
    "Sanco": filter_sanco,
}

def process_file(file_path: str, writer) -> None:
    """Process a single file: extract, filter, and write to workbook."""
    company = detect_company_from_filename(file_path)
    ext = os.path.splitext(file_path)[1].lower()
    try:
        if ext == ".pdf":
            df = extract_pdf_to_dataframe(file_path)
        elif ext in [".xls", ".xlsx"]:
            df = pd.read_excel(file_path, header=None)
        else:
            return
        if df.empty:
            logging.warning(f"No data extracted from {file_path}")
            return
        if DEBUG_SAVE_RAW:
            raw_path = os.path.splitext(file_path)[0] + "_RAW.xlsx"
            df.to_excel(raw_path, index=False, header=False)
            logging.info(f"Saved raw extraction to {raw_path}")
        if company in COMPANY_FILTERS:
            filtered = COMPANY_FILTERS[company](df)
            if filtered.empty:
                logging.warning(f"No filtered data for {company} in {file_path}")
                return
            logging.info(f"Extracted columns for {company}: {os.path.basename(file_path)}")
            filtered.to_excel(writer, sheet_name=company, index=False, header=False)
        else:
            logging.info(f"No filter defined for {company}, skipping {file_path}")
    except Exception as e:
        logging.error(f"Failed processing {file_path}: {e}")
