import PyPDF2
import pandas as pd
import re


__all__ = [
    "execute"
]


def data_parse_with_regex(input_string):
    try:
        if 'of 552' not in input_string:
            date_matches = re.findall(r'\d{1,2}/[A-Za-z]{3}/\d{4}', input_string)
            last_date = date_matches[-1]
            split_parts = input_string.rsplit(last_date, 1)
            first_part = split_parts[0].strip()
            first_part_list = first_part.split(" ")
            first_part_list.append(last_date)
            
            last_part = split_parts[1].strip().split(" ")
            company_name_list = last_part[0:-6]
            company_name = " ".join(company_name_list)
            first_part_list.append(company_name)
            remaining_data = last_part[-6:]
            first_part_list.extend(remaining_data)
            return first_part_list
        else:
            print("its a page string",input_string)
            return []
    except Exception as e:
        print(e)
        return []



def extract_table_from_pdf(pdf_path):
    table_data = []
    with open(pdf_path, 'rb') as pdf_file:
        pdf_reader = PyPDF2.PdfReader(pdf_file)
        for page in pdf_reader.pages:
            text = page.extract_text()
            lines = text.split('\n')
            for line in lines:
                if "Sr No." not in line and "Encashment" not in line and "Political" not in line and 'Branch' not in line and 'Teller' not in line:
                    if len(line) >= 6:  # Assuming at least 6 characters for each line
                        if len(data_parse_with_regex(line)) > 0:
                            table_data.append(data_parse_with_regex(line))
            
    return table_data

def save_to_excel(table_data, output_path):
    try:
        df = pd.DataFrame(table_data, columns=['Sr No.', 'Date Of Encashment', 'Name of Political party', 'Account Number', 
                                               'Prefix', 'Bond Number', 'Denominations', 'Pay Branch Code', 'Pay Teller'])
        df.to_excel(output_path, index=False)
    except Exception as e:
        print("some error occured converting to xlsx")
    print("File conversion complete")

def convert_user_pdf_to_excel(pdf_path, output_path):
    table_data = extract_table_from_pdf(pdf_path)
    save_to_excel(table_data, output_path)


def convert_user(output_file):
    print("Staring to convert bond buyer pdf to xlsx")
    pdf_path = 'bond_user.pdf'  # Path to your PDF file
    output_path = output_file
    convert_user_pdf_to_excel(pdf_path, output_path)
    print("Completed Conversion of PDf  Excel")

