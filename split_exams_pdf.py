import sys
from operator import itemgetter
from PyPDF2 import PdfReader, PdfWriter
import pandas
from pathlib import Path


def print_usage():
    print("Usage: split_exams_pdf.py <path_to_pdf> <path_to_spreadsheet> <spreadsheet_tab_index>")


def exit_error_with_usage():
    print_usage()
    exit(1)


def exit_error_with_msg(msg: str):
    print(msg)
    exit(1)


def read_cmd_args():
    if len(sys.argv) < 4:
        exit_error_with_usage()

    try:
        tab_idx = int(sys.argv[3])
    except ValueError:
        tab_idx = -1  # should never matter; exiting on next line
        exit_error_with_usage()

    return Path(sys.argv[1]), Path(sys.argv[2]), tab_idx


def read_pdf_info(spreadsheet_path: Path, tab_idx: int):
    spreadsheet = pandas.read_excel(spreadsheet_path, header=2, sheet_name=tab_idx)
    metadata = spreadsheet.loc[spreadsheet["Seite"].notnull()][["Seite", "Name", "Vorname"]].values.tolist()
    metadata.sort(key=itemgetter(0))

    pdf_info = []
    for index, line in enumerate(metadata):
        seite, name, vorname = line
        if index < len(metadata) - 1:
            next_row = metadata[index + 1]
            next_seite = next_row[0]
            end_before_page = int(next_seite)
        else:
            end_before_page = None

        pdf_info.append({
            "first_name": vorname,
            "last_name": name,
            "start_page": int(seite),
            "end_before_page": end_before_page
        })

    return pdf_info


def export_pdfs(exams_pdf_path: Path, pdf_info):
    out_dir_path = exams_pdf_path.parent / exams_pdf_path.stem
    if out_dir_path.exists():
        exit_error_with_msg(f"Output directory path already exists: {out_dir_path}")
    out_dir_path.mkdir()

    infile = PdfReader(exams_pdf_path)
    end_of_file = len(infile.pages) + 1

    for student in pdf_info:
        output = PdfWriter()
        end_before_page = student["end_before_page"] if student["end_before_page"] is not None else end_of_file
        for page in range(student["start_page"], end_before_page):
            output.add_page(infile.pages[page - 1])

        out_file_path = out_dir_path / f"{student['first_name']}_{student['last_name']}.pdf"
        print(f"Writing individual exam PDF to ${out_file_path}")
        with open(out_file_path, "wb") as outfile:
            output.write(outfile)


def main():
    exams_pdf_path, spreadsheet_path, tab_idx = read_cmd_args()

    if not exams_pdf_path.exists():
        exit_error_with_msg(f"Exams PDF not found at path {exams_pdf_path}")

    if not spreadsheet_path.exists():
        exit_error_with_msg(f"Metadata spreadsheet not found at path {spreadsheet_path}")

    pdf_info = read_pdf_info(spreadsheet_path, tab_idx)
    export_pdfs(exams_pdf_path, pdf_info)


if __name__ == "__main__":
    main()
