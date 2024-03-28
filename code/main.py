import pandas as pd
import os
from bond_buyer_file_converter import convert_buyer
from bond_user_file_converter import convert_user


def merge_both_excel_sheet(buyer_file, user_file):
    # Read the Excel sheets into DataFrames
    df1 = pd.read_excel(buyer_file)
    df2 = pd.read_excel(user_file)

    # Merge the DataFrames based on two common columns
    merged_df = pd.merge(df1, df2, on=['Prefix', 'Bond Number'])

    # Write the merged DataFrame to a new Excel file
    merged_df.to_excel('xlsx/merged_buyer_and_user_test.xlsx', index=False)


def converter(buyer_output, user_output):
    try:
        convert_buyer(buyer_output)
        convert_user(user_output)
    except Exception as e:
        print("Some error Occured")

def create_folder_if_not(folder_name):
    # Define the folder path

    # Create the folder if it doesn't exist
    if not os.path.exists(folder_name):
        os.makedirs(folder_name)


def build_markdown():
   # Read the merged Excel data into a Pandas DataFrame
    merged_df = pd.read_excel('xlsx/merged_buyer_and_user.xlsx')

    # Extract the year from the relevant date column
    merged_df['Year'] = pd.to_datetime(merged_df['Date of Purchase']).dt.year
    merged_df['Denominations_x'] = merged_df['Denominations_x'].str.replace(',', '').astype(float)
    grouped_df = merged_df.groupby(['Year', 'Name of the Purchaser']).agg({'Denominations_x': 'sum'}).reset_index()
    grouped_df.rename(columns={'Denominations_x': 'Total Amount of Bond Bought'}, inplace=True)
    grouped_df['Total Amount of Bond Bought'] = grouped_df['Total Amount of Bond Bought'].apply(lambda x: '{:.2f}'.format(x))
    
    folder_name = "content"
    create_folder_if_not(folder_name)

    # Write the table to a Markdown file in the content folder
    markdown_file_path = os.path.join(folder_name, 'most_frequent_purchaser.md')
    grouped_df.to_html(markdown_file_path, index=False)
        
    # # Table 2: Mean Purchaser Amount by Year, Sorted Descendingly
    # Extract the year from the relevant date column
    merged_df['Year'] = pd.to_datetime(merged_df['Date Of Encashment']).dt.year
    merged_df['Denominations_y'] = merged_df['Denominations_y'].str.replace(',', '').astype(float)
    grouped_df = merged_df.groupby(['Year', 'Name of Political party']).agg({'Denominations_y': 'sum'}).reset_index()
    grouped_df.rename(columns={'Denominations_y': 'Total amount of bond encahsed'}, inplace=True)
    grouped_df['Total amount of bond encahsed'] = grouped_df['Total amount of bond encahsed'].apply(lambda x: '{:.2f}'.format(x))
    
    
    # Write the table to a Markdown file in the content folder
    markdown_file_path = os.path.join(folder_name, 'most_frequent_encashment.md')
    grouped_df.to_html(markdown_file_path, index=False)

    

    
def main():
    buyer = "xlsx/bond_buyer_test.xlsx"
    user = "xlsx/bond_user_test.xlsx"
    print("Converting file from pdf to excel to do some magic")
    converter(buyer, user)
    print("File conversion complete........")
    print("Merging both the files On Prefix and Bond Number ..................")
    merge_both_excel_sheet(buyer, user)
    print("File merger complete.....................................")
    print("Check file merged_buyer_and_user to see merged data")
    




if __name__ == "__main__":
    main()
    build_markdown()