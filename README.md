# Data-Parsing-in-excel

This Python script helps you extract and organize vessel and cargo information from a specific Excel report format into a clean, structured table.

## 🚀 How it Works

Imagine you have an Excel file about port activities (like ships arriving or already at berth, and what cargo they carry). This script reads that file, understands the different sections (like "Expected" or "At Berth"), identifies cargo types (like "Crude oil" or "Diesel"), and pulls out key details like vessel names, dates, and quantities.

It then takes all this extracted information and saves it into a brand new, easy-to-read Excel file.

## ✨ Features

* **Parses specific Excel layouts:** Designed to understand sections like "Expected" and "At Berth" and cargo types.
* **Extracts key data:** Grabs vessel names, ETA/Berthed dates, cargo types, and quantities.
* **Cleans and formats data:** Ensures dates are consistent and quantities are presented clearly.
* **Generates a new, organized Excel file:** Outputs a structured table ready for analysis.

## 📁 Files in this Repository

* `Data_Parsing.ipynb`: The main Python script that does all the parsing and saving.
* `Awesome port.xlsx` (Example Input): This is the *type* of Excel file the script expects to read. You'll need an Excel file with a similar structure for the script to work correctly.

## 🛠️ How to Use

1.  **Save your input file:** Make sure your Excel report file is named `Awesome port.xlsx` and is in the same directory as the script. If your file has a different name, you'll need to update the `input_file` variable in the script:
    ```python
    input_file = '/content/Awesome port.xlsx' # Change this if your file name is different
    ```
    (Note: The `/content/` part is typical for Google Colab environments. If you're running this locally, it might just be `Awesome port.xlsx` or a full path like `C:/Users/YourUser/Documents/Awesome port.xlsx`)

2.  **Run the script:**
    If you are using Jupyter Notebook or Google Colab, simply run all the cells in the `Data_Parsing.ipynb` file.

    If you want to run it as a standalone Python script (after converting or copying the code into a `.py` file, e.g., `parse_port_data.py`), open your terminal or command prompt, navigate to the directory where you saved the file, and run:
    ```bash
    python Data_Parsing.ipynb # Or python parse_port_data.py
    ```

3.  **Find your output:** A new Excel file named `Parsed_Output.xlsx` will be created in the same directory, containing your organized data.

## 📝 Requirements

This script uses the `openpyxl` library. You can install it using pip:

```bash
pip install openpyxl

4. ##Conclusion

This tool streamlines the process of extracting and organizing port data from Excel files, saving you time and effort. It provides a clear, structured output that can be easily used for further analysis or reporting.

Found this useful, star and fork this repository.
