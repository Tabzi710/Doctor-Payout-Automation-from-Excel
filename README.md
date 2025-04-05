Python script to extract doctor-wise information from consolidated Excel files and calculate daily revenue.

Installation

1. Ensure you have Python installed on your system.

2. Install the required packages:
   ```
   pip install pandas openpyxl
   ```

3. Clone or download this repository to your local machine.

Usage

1. Run the script:
   ```
   EMD.py
   ```

2. When prompted:
   - Enter the full path to your Excel file
   - Enter the output directory path where all generated files will be saved (press Enter to use the current directory)

3. The script will generate the following Excel files in the output path.

The script will attempt to match columns based on similar names, so slight variations in column names should still work.
