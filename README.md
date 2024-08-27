# excel_functions

This contains a script that allows users to copy values from a source Excel file to a template Excel file, saving the result as a new output file. It utilizes the `tkinter` library for GUI dialogs and `openpyxl` for Excel file manipulation.

## Prerequisites

Before running the script, ensure you have the following installed:
- Python 3.x
- `tkinter` (usually included with Python)
- `openpyxl` library

You can install `openpyxl` using pip if it is not already installed:
```bash
pip install openpyxl
```

## Usage Instructions


1. **Clone/Download the Repository:**

   - If you do not have Git installed:
      Go to the repository page on GitHub.
      Click on the "Code" button.
      Select "Download ZIP".
      Extract the ZIP file to your desired location.
      Open a terminal or command prompt and navigate to the extracted directory.
     
   - OR use git
     
   ```bash
   git clone <repository-url>
   cd <repository-directory>
   ```

## to_template.py ##################################

**Prepare the Excel Files:**

Place the Excel files you want to use (source and template) in the same directory as the script.
Ensure the files are in `.xlsx` format.

### 1. Run the Script:

Execute the Python script using the command line:

```bash
python to_template.py
```

### 2. Select the Files:

- The script will display a list of all `.xlsx` files in the directory.
- You will be prompted to select the **source file** and the **template file** by entering the corresponding numbers from the list.

### 3. Define the Range:

- After selecting the files, you will be asked to define the maximum column and row to copy:
  - **Column**: Enter the maximum column number to copy (e.g., `17` for column `Q`).
  - **Row**: Enter the maximum row number to copy (e.g., `100`).

### 4. Output File Name:

- You will be prompted to enter the name for the output file.
- The output file will be saved in the same directory as the script with the specified name.

### 5. Check the Results:

- The script will copy the data from the source file to the template within the specified range.
- The modified template will be saved as a new Excel file with the name you provided.

## Example Workflow

1. **Prepare Files**: Place `data.xlsx` (source) and `template.xlsx` in the same directory as the script.
2. **Run Script**: Execute the script using `python excel_data_copier.py`.
3. **Select Files**: Choose `data.xlsx` as the source and `template.xlsx` as the template.
4. **Define Range**: Enter `17` for columns (up to column `Q`) and `100` for rows (up to row `100`).
5. **Output File Name**: Enter `final_output` when prompted. The file `final_output.xlsx` will be created.

## Additional Information

- **Default File Handling**: If no file name is provided, the output file will default to `output.xlsx`.
- **Column and Row Limits**: The script allows copying up to the maximum possible columns (256) and rows (1,048,576) supported by Excel.

## excel_cleaner.py #############################


This script is designed to clean and process Excel files by filtering out non-English characters and replacing specific non-ASCII characters with their ASCII equivalents. The cleaned data is saved in a new Excel file.

## Prerequisites

Ensure you have Python installed on your system. Additionally, you will need the `openpyxl` library to work with Excel files. You can install this library using pip:

```bash
pip install openpyxl
```

## Usage

### 1. Prepare the Excel File:

1. Place the Excel file(s) you want to process in the same directory as the script.
2. The script will automatically list all `.xlsx` files in the directory.

### 2. Run the Script:

Execute the Python script using the command line:

```bash
python excel_cleaner.py
```

### 3. Select the File:

- The script will prompt you to select the Excel file to clean from a list of available files in the directory.
- Use the number corresponding to your desired file to select it.

### 4. Cleaning Process:

- The script processes the selected file by filtering out non-English characters by default.
- Additional functions for replacing specific characters (`ø`, `ö`, `å`, `ä`, `º`, `≤`, etc.) with their ASCII equivalents are available. These can be activated by uncommenting the relevant lines in the script.

### 5. Check the Results:

- After processing, the script will save the cleaned content to a new Excel file. The new file is named by appending `_cleaned` to the original filename (e.g., `original_file_cleaned.xlsx`).
- The script will attempt to open the new file to verify that it was saved correctly. If successful, a confirmation message will be printed in the console.

## Example Output

For an Excel file named `output.xlsx`, the cleaned file will be saved as `output_cleaned.xlsx` in the same directory.

## Additional Information

- **Filtering Non-English Characters**: By default, the script removes all non-ASCII characters.
- **Character Replacements**: 
  - `ø` can be replaced with `phase`
  - `ö` can be replaced with `oe`
  - `å` can be replaced with `aa`
  - `ä` can be replaced with `ae`
  - `º` can be replaced with `degrees`
  - `≤` can be replaced with `<=`
  
  To enable these replacements, uncomment the relevant lines in the script.


