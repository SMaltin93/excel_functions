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

1. **Clone the Repository:**
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
   - 

6. **Output File:**
   - The script will create a new Excel file named `output.xlsx` in the same directory with the copied data.
   - Remember the source and template files-numbers that is showing from the prompt and then enter the corresponding number.

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



