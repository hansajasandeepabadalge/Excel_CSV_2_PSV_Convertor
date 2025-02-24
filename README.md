# CSV/Excel to PSV Converter

## Description
This is a simple GUI-based application built with Python and Tkinter that allows users to convert CSV and Excel files into pipe-separated value (PSV) files. The converted files are saved in the `Documents/PSV_Files` directory.

## Features
- Supports both `.csv` and `.xlsx/.xls` file formats.
- Converts data to PSV format (pipe `|` separated values).
- Automatically creates a `PSV_Files` folder in the `Documents` directory.
- Provides a simple GUI for file selection and conversion.

## Prerequisites
Ensure you have Python installed on your system. You also need the following Python packages:

- `tkinter` (Built-in with Python)
- `pandas`

If `pandas` is not installed, you can install it using:
```sh
pip install pandas
```

## Installation
1. Clone this repository or download the script.
2. Ensure you have Python installed (version 3.x recommended).
3. Install the required dependencies.

## Usage
1. Run the script:
   ```sh
   python csv2psv.py
   ```
2. Click the **Browse** button to select a CSV or Excel file.
3. Click **Convert** to convert the file into PSV format.
4. The converted file will be saved in `Documents/PSV_Files`.

## Error Handling
- If no file is selected, an error message will be displayed.
- If an invalid file format is chosen, the program will notify the user.
- Errors during conversion will be displayed in the status label.

## Future Enhancements
- Support for additional file formats.
- Drag-and-drop functionality for file selection.
- Option to choose output directory.

## License
This project is licensed under the MIT License.

