# excel_functions

## to_template.py

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

2. **Prepare Your Files:**
   - Place your source and template Excel files in the same directory as the script. Ensure these files have the `.xlsx` extension.

3. **Run the Script:**
   ```bash
   python script.py
   ```

4. **Input the Maximum Column and Row:**
   - The script will prompt you to enter the maximum column and row to copy. Enter these values based on your data range.

5. **Select the Source and Template Files:**
   - The script will display a list of Excel files found in the directory.
   - Remember the source and template files-numbers that is showing from the prompt and then enter the corresponding number.

6. **Output File:**
   - The script will create a new Excel file named `output.xlsx` in the same directory with the copied data.


## Example

```bash
python to_template.py
```

1. **Input Maximum Column and Row:**
   - Enter `17` for the maximum column (e.g., column Q).
   - Enter `100` for the maximum row.

2. **Select Files:**
   - Choose the source file number from the displayed list.
   - Choose the template file number from the displayed list.

3. **Output:**
   - The script generates an `output.xlsx` file with the copied data.




## excel_cleaner.py


## Prerequisites

Ensure you have Python installed on your system along with the required libraries. You can install the necessary Python libraries using `pip`:

```bash
pip install openpyxl
```

## Usage

1. **Prepare the Excel File:**
   - Place the Excel file you want to process in the same directory as the script.
   - Ensure the file is named `output.xlsx`, or modify the script to match the name of your Excel file.

2. **Run the Script:**
   - Execute the Python script using the command line:

   ```bash
   python excel_cleaner.py
   ```

   - The script will load the `output.xlsx` file, filter out non-English characters, and save the modified content to a new file named `Modified.xlsx`.

3. **Check the Results:**
   - The new Excel file (`Modified.xlsx`) will be saved in the same directory.
   - The script will attempt to open the new file to verify it was saved correctly. If successful, a confirmation message will be printed.


