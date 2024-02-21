import openpyxl
import csv
import os

def converter(input_excel_file, output_csv_file):

    wb = openpyxl.load_workbook(input_excel_file)
    sheet = wb.active

    with open(output_csv_file, 'w', newline='') as csvfile:
        csv_writer = csv.writer(csvfile)

        for row in sheet.iter_rows(min_row=1, values_only=True):
            csv_writer.writerow(row)

    print(f"Conversion complete! check '{output_csv_file}' ")

input_excel_file = 'archive/quarters.xlsx'
output_csv_file = 'data/quarters.csv'
converter(input_excel_file, output_csv_file)

def remove_null_rows_and_columns(input_csv_file, output_csv_file):
    with open(input_csv_file, 'r') as fin, open(output_csv_file, 'w', newline='') as fout:
        csv_reader = csv.reader(fin)
        csv_writer = csv.writer(fout)

        # Transpose the matrix (swap rows with columns) to process columns as rows
        transposed_rows = zip(*csv_reader)

        # Only write columns with non-null values
        non_null_columns = [column for column in transposed_rows if any(column)]
        transposed_non_null_rows = zip(*non_null_columns)

        for row in transposed_non_null_rows:
            if any(row):
                csv_writer.writerow(row)

    print(f"Null rows and columns removed from '{input_csv_file}' and saved to '{output_csv_file}'.")

# Example usage for removing null rows and columns
input_csv_file = 'data/quarters.csv'
output_csv_file = 'data/quarters_nn.csv'
remove_null_rows_and_columns(input_csv_file, output_csv_file)

def delete_file(file_path):
    try:
        os.remove(file_path)
        print(f"File '{file_path}' deleted successfully.")
    except FileNotFoundError:
        print(f"File '{file_path}' not found.")
    except Exception as e:
        print(f"Error deleting file '{file_path}': {e}")

# Example usage
file_to_delete = 'data/quarters.csv'
delete_file(file_to_delete)
