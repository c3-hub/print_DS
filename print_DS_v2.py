import xlwings as xw
import re
import os


def find_page_breaks(sheet):
    # Get the Excel application object
    xl_app = xw.apps.active

    # Get the print area of the sheet
    print_area = sheet.api.PageSetup.PrintArea
    # Get the range of the print area
    print_range = sheet.range(print_area)

    # Select the print range to work with page breaks
    print_range.select()
    xl_app.api.ActiveWindow.View = 2  # Page Break Preview

    # Get page breaks
    h_page_breaks = sheet.api.HPageBreaks
    v_page_breaks = sheet.api.VPageBreaks

    # Return page break locations
    h_positions = [pb.Location for pb in h_page_breaks]
    v_positions = [pb.Location for pb in v_page_breaks]

    return h_positions, v_positions


def main():
    # Connect to the active Excel instance
    wb = xw.books.active
    # Get the active worksheet
    sheet = wb.sheets.active

    # Get the directory path of the active Excel file
    excel_file_path = wb.fullname
    excel_dir = os.path.dirname(excel_file_path)

    # Find page breaks
    h_page_breaks, v_page_breaks = find_page_breaks(sheet)

    # Add the start and end positions to handle the first and last pages
    if len(h_page_breaks) == 0:
        print("Vertical Page Breaks are not defined, return to Excel!")
    else:
        h_page_breaks = [0] + [pb.Row for pb in h_page_breaks]

    v_page_breaks = [0] + [pb.Column for pb in v_page_breaks] + [sheet.used_range.last_cell.column]

    print("The boundaries of the workspace are set:")
    print(f"    HPageBreaks rows: {h_page_breaks}")
    print(f"    HPageBreaks count: {len(h_page_breaks)}")
    print(f"    VPageBreaks columns: {v_page_breaks}")
    print(f"    VPageBreaks count: {len(v_page_breaks)}" + '\n')

    # Iterate through each page
    for i in range(len(h_page_breaks) - 1):
        for j in range(len(v_page_breaks) - 1):
            start_row = h_page_breaks[i]
            end_row = h_page_breaks[i + 1]
            start_col = v_page_breaks[j]
            end_col = v_page_breaks[j + 1]

            if j == 0:
                # Get the range of the first page
                page_range = sheet.range((start_row + 1, start_col + 1), (end_row, end_col - 1))

            else:
                # Get the range of the current page
                page_range = sheet.range((start_row + 1, start_col), (end_row, end_col - 1))

            # Iterate through each cell in the third column
            for cell in page_range.columns[2]:
                # Check if the cell matches the specified pattern
                if re.match(r'C3\.\w{2,3}\d{3,4}', str(cell.value)):
                    # Extract the name from the cell
                    name = str(cell.value)
                    # Define the PDF filename
                    pdf_filename = f"DS_{name}.pdf"  # Add page number to avoid overwriting
                    # Define the full path to save the PDF
                    pdf_path = os.path.join(excel_dir, pdf_filename)
                    # Print the page to PDF
                    page_range.api.ExportAsFixedFormat(0, pdf_path)
                    print(f"Page printed: {pdf_filename}")
                    break

        print('\n' + "Printing is over!")


if __name__ == "__main__":
    main()
