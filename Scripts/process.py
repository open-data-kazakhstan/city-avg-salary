import openpyxl
import csv
import os
from datapackage import Package


def converter(input_excel_file, output_csv_file):
    wb = openpyxl.load_workbook(input_excel_file)
    sheet = wb.active

    with open(output_csv_file, 'w', newline='') as csvfile:
        csv_writer = csv.writer(csvfile)

        for row in sheet.iter_rows(min_row=1, values_only=True):
            csv_writer.writerow(row)

    print(f"Conversion complete! check '{output_csv_file}' ")


input_excel_file = 'archive/pg16 optRK2022.xlsx'
first_out = 'data/avg_salary.csv'
converter(input_excel_file, first_out)


def renamer(input_file, output_file):
    column_mapping = {
        'Регионы': 'regions',
        'Всего персонала': 'avg_salary',
        'руководители и государственные служащие': 'managers_&_civil_servants',
        'специалисты-профессионалы': 'qualified',
        'специалисты-техники и иной вспомогательный профессиональный персонал': 'tech_personnel',
        'служащие в области администрирования': 'admin_emplyees',
        'работники сферы услуг и продаж': 'service_and_sales',
        'фермеры и рабочие сельского и лесного хозяйства, рыбоводства и рыболовства': 'agriculture',
        'рабочие промышленности, строительства, транспорта и других родственных занятий': 'Engineers_transport',
        'операторы производственного оборудования, сборщики и водители': 'assemblers_and_drivers',
        'неквалифицированные рабочие': 'unqualified',
        'работники, не входящие в другие группы ': 'not_members_of_other_groups'
    }

    with open(input_file, 'r', encoding='utf-8') as infile, open(output_file, 'w', encoding='utf-8', newline='') as outfile:
        reader = csv.reader(infile)
        writer = csv.writer(outfile)

        header = next(reader)
        new_header = [column_mapping.get(col, col) for col in header]
        writer.writerow(new_header)

        for row in reader:
            writer.writerow(row)


def text_remover(input_file, output_file):
    with open(input_file, 'r', encoding='utf-8') as infile, open(output_file, 'w', encoding='utf-8', newline='') as outfile:

        skip_line = False

        for line in infile:
            if "16. Среднемесячная заработная" in line or "тенге" in line:
                skip_line = True
                continue
            if not skip_line:
                outfile.write(line)
            skip_line = False


def delete_file(file_path):
    try:
        os.remove(file_path)
        print(f"File '{file_path}' deleted successfully.")
    except FileNotFoundError:
        print(f"File '{file_path}' not found.")
    except Exception as e:
        print(f"Error deleting file '{file_path}': {e}")


def create_datapackage(data_file, metadata_file='datapackage.json'):
    package = Package()

    resource = package.add_resource({"path": "data/avg_salary_.csv",
            "profile": "data-resource",
            "name": "avg_salary_",
            "format": "csv",
            "mediatype": "text/csv",
            "encoding": "utf-8",
            "schema": {
                "fields":[
            {
                "name": "regions",
                "type": "string",
                "format": "default"
            },
                    {
                "name": "avg_salary",
                "type": "integer",
                "format": "default"
            },
                    {
                "name": "managers_&_civil_servants",
                "type": "integer",
                "format": "default"
            },
                    {
                "name": "qualified",
                "type": "integer",
                "format": "default"
            },
                    {
                "name": "tech_personnel",
                "type": "integer",
                "format": "default"
            },
                    {
                "name": "admin_emplyees",
                "type": "integer",
                "format": "default"
            },
                    {
                "name": "service_and_sales",
                "type": "integer",
                "format": "default"
            },
                    {
                "name": "agriculture",
                "type": "integer",
                "format": "default"
            },
                    {
                "name": "transport",
                "type": "integer",
                "format": "default"
            },
                    {
                "name": "assemblers_and_drivers",
                "type": "integer",
                "format": "default"
            },{
                "name": "unqualified",
                "type": "integer",
                "format": "default"
            },{
                "name": "not_members_of_other_groups",
                "type": "integer",
                "format": "default"
            }
                ]}})

    resource.infer()

    package.save(metadata_file)
    print(f"Datapackage created! Check '{metadata_file}' ")


second_out = 'data/avg_salary_.csv'
renamer(first_out, second_out)
delete_file(first_out)
create_datapackage(second_out)