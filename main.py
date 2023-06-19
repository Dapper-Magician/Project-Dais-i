# Importing necessary libraries
import os
from datetime import datetime

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

# Specify the folder path containing the Excel files
file_path = "C:/Users/moons/OneDrive/Desktop/CMAA-Daisi/Bot_Cleaning"


class ExcelScrapingChatbot:

    def __init__(self):
        self.excel_file = None
        self.data_frames = {}

    def upload_excel(self, file_path):
        try:
            self.excel_file = pd.ExcelFile(file_path)
            self.data_frames = {sheet_name: None for sheet_name in self.excel_file.sheet_names}
            return "Excel file uploaded successfully."
        except FileNotFoundError:
            return "File not found. Please provide a valid file path."
        except Exception as e:
            return f"An error occurred while uploading the Excel file: {str(e)}"

    def process_excel_files(folder):
        # Get a list of all files in the folder
        files = os.listdir(folder)

        # Iterate over the files
        for file in files:
            # Check if the file is an Excel file
            if file.endswith('.xlsx') or file.endswith('.xls'):
                # Create the full file path
                file_path = os.path.join(folder, file)

                # Perform the desired operations on the Excel file
    process_excel_file(file_path)


    def process_excel_file(Daisi_CleanSheets.xlsx):
    # Create an instance of the ExcelScrapingChatbot
        scraping_chatbot = ExcelScrapingChatbot()

        # Retrieve the data from the Excel file using the scraping chatbot
        sheet_name, column_name, start_row, end_row = gather_user_input()  # Modify this as needed
        retrieved_data = scraping_chatbot.retrieve_data(sheet_name, column_name, start_row, end_row)

        # Create an instance of the ExcelCleaningChatbot
        cleaning_chatbot = ExcelCleaningChatbot()

        # Perform cleaning and organization of the scraped data using the cleaning chatbot
        cleaned_data_frame = cleaning_chatbot.clean.excel()
        organized_data_frame = cleaning_chatbot.organize_excel()

        # Export the cleaned and organized data to a new Excel file
        output_file = "Daisi_CleanSheets.xlsx"  # Modify this as needed
        output_folder = "C:/Users/moons/OneDrive/Desktop/CMAA-Daisi/Bot_Cleaning"  # Modify this as needed
        cleaning_chatbot.export_excel(organized_data_frame, output_file, output_folder)


# Call the function to process the Excel files in the folder
    process_excel_files(file_path)


    def retrieve_data(self, sheet_name, columns, start_row, end_row):
        data_frame = pd.read_excel(self.excel_file, sheet_name=sheet_name)
        # Filter rows based on start_row and end_row
        data_frame = data_frame.iloc[start_row - 1:end_row, :]
        if sheet_name not in self.excel_file.sheet_names:
            return f"The sheet '{sheet_name}' does not exist in the Excel file."

        if self.data_frames[sheet_name] is None:
            self.data_frames[sheet_name] = self.excel_file.parse(sheet_name)

        # Retrieve the specified columns or all columns if column_names is None
        if columns is not None:
            data_frame = data_frame[columns]

            data_frame = self.data_frames[sheet_name]

        if not columns:
            return "No columns specified. Please provide the column names."

        if start_row is not None and end_row is not None:
            data_frame = data_frame.loc[start_row:end_row, columns]
        else:
            data_frame = data_frame.loc[:, columns]

        return data_frame

    def generate_excel(self, output_file):
        writer = pd.ExcelWriter(output_file, engine='xlsxwriter')

        for sheet_name, data_frame in self.data_frames.items():
            if data_frame is not None:
                data_frame.to_excel(writer, sheet_name=sheet_name, index=False)

        writer.save()
        return "Excel file generated successfully."

    def handle_user_query(self, query):
        if query == "upload":
            return "Please provide the file path for the Excel file."

        if query.startswith("retrieve"):
            _, sheet_name, column_name = query.split(" ", 2)
            data = self.retrieve_data(sheet_name, column_name)

            if isinstance(data, list):
                return data
            else:
                return data

        return "I'm sorry, I couldn't understand your query. Please try again."


class ExcelCleaningChatbot:
    def __init__(self):
        self.excel_file = None
        self.data_frame = None

    def upload_excel(self, file_path):
        try:
            self.data_frame = pd.read_excel(file_path)
            self.excel_file = file_path
            return "Excel file uploaded successfully."
        except FileNotFoundError:
            return "File not found. Please provide a valid file path."
        except Exception as e:
            return f"An error occurred while uploading the Excel file: {str(e)}"

    def clean_excel(self):
        # Perform cleaning operations on the Excel sheet
        # (Add your specific cleaning logic here)
        cleaned_data_frame = self.data_frame

        return cleaned_data_frame

    def organize_excel(self):
        # Perform organizing operations on the cleaned Excel sheet
        # (Add your specific organizing logic here)
        organized_data_frame = self.data_frame

        return organized_data_frame

    def export_excel(self, data_frame, output_file, output_folder):
        try:
            workbook = Workbook()
            sheet = workbook.active

            # Write column headers
            for col_num, column_name in enumerate(data_frame.columns, 1):
                cell = sheet.cell(row=1, column=col_num, value=column_name)
                cell.font = Font(bold=True)

            # Write data rows
            for row_num, row in enumerate(data_frame.itertuples(), 2):
                for col_num, value in enumerate(row[1:], 1):
                    sheet.cell(row=row_num, column=col_num, value=value)

            # Set column widths and alignment
            for column_cells in sheet.columns:
                max_length = 0
                column = column_cells[0].column_letter
                for cell in column_cells:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                adjusted_width = (max_length + 2) * 1.2
                sheet.column_dimensions[column].width = adjusted_width
                for cell in column_cells:
                    cell.alignment = Alignment(wrap_text=True)

            # Generate a unique file name based on the current date and time
            timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            file_name = f"{output_file}_{timestamp}.xlsx"

            # Create the output folder if it doesn't exist
            if not os.path.exists(output_folder):
                os.makedirs(output_folder)

            # Save the workbook in the output folder
            output_path = os.path.join(output_folder, file_name)
            workbook.save(output_path)

            return "Excel file exported successfully."

        except Exception as e:
            return f"An error occurred while exporting the Excel file: {str(e)}"


def gather_user_input():
    sheet_name = input("Sheet Name: ")
    column_name = input("Column Name: ")
    start_row = int(input("Start Row: "))
    end_row = int(input("End Row: "))

    return sheet_name, column_name, start_row, end_row


# Main function
def main():
    scraping_chatbot = ExcelScrapingChatbot()
    cleaning_chatbot = ExcelCleaningChatbot()

    print("Welcome to the Excel Toolbox Chatbot!")
    print("Which action would you like to perform?")
    print("1. Scraping and exporting data to Excel.")
    print("2. Cleaning and organizing an Excel sheet.")

    bot_choice = input("> ")

    if bot_choice == "1":
        sheet_name, column_name, start_row, end_row = gather_user_input()

        retrieved_data = scraping_chatbot.retrieve_data(sheet_name, [column_name], start_row, end_row)
        print("Retrieved Data:")
        print(retrieved_data)

    elif bot_choice == "2":
        print("Enter the output file name for the exported Excel file: ")
        output_file = input("> ")
        output_folder = input("Enter the output folder path for the exported Excel file: ")

        organized_data_frame = cleaning_chatbot.organize_excel()
        export_response = cleaning_chatbot.export_excel(organized_data_frame, output_file, output_folder)
        print(export_response)

    else:
        print("Invalid choice. Please try again.")


# Entry point of the program
if __name__ == "__main__":
    main()
