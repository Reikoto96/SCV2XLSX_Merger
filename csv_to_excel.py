import os  # Import the os module for interacting with the operating system
import pandas as pd  # Import the pandas library for data manipulation and analysis
import shutil  # Import the shutil module for file operations
import openpyxl  # Import the openpyxl module for working with Excel files

def main():
    # Define the current working directory
    cwd = os.getcwd()

    # Define the path for the old directory
    old_dir = os.path.join(cwd, "old")  # Join the current working directory with "old" folder name
    if not os.path.exists(old_dir):  # Check if the "old" directory does not exist
        os.makedirs(old_dir)  # Create the "old" directory

    # Find all CSV files in the directory
    csv_files = [f for f in os.listdir(cwd) if f.endswith('.csv')]  # List all files in the current directory ending with '.csv'

    # List to hold the dataframes
    dataframes = []

    for csv_file in csv_files:  # Iterate over each CSV file found
        # Read each CSV file into a dataframe
        df = pd.read_csv(os.path.join(cwd, csv_file))  # Read the CSV file into a pandas DataFrame
        dataframes.append(df)  # Append the DataFrame to the list

        # Move the CSV file to the old directory
        shutil.move(os.path.join(cwd, csv_file), os.path.join(old_dir, csv_file))  # Move the CSV file to the "old" directory

    # Concatenate all dataframes
    if dataframes:  # Check if there are any dataframes in the list
        merged_df = pd.concat(dataframes, ignore_index=True)  # Concatenate all dataframes into a single dataframe

        # Define the path for the output Excel file
        output_file = os.path.join(cwd, "merged_output.xlsx")  # Define the path for the output Excel file

        # Write the merged dataframe to an Excel file
        merged_df.to_excel(output_file, index=False)  # Write the merged dataframe to an Excel file without row index

        # Open the output file with Excel
        os.system(f'start "" "C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE" "{output_file}"')  # Open the output Excel file with Excel application
    else:
        print("No CSV files found in the directory.")  # Print a message if no CSV files are found in the directory

if __name__ == "__main__":
    main()  # Call the main function if the script is executed directly
