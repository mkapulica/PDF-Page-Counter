# PDF Page Counter

**PDF Page Counter** is a Python-based desktop application designed to count the number of pages in PDF files located in a specific directory. The application offers the following features:

- **Graphical User Interface**: User-friendly GUI created with Tkinter.
- **Directory Selection**: Choose a directory and optionally include subfolders for PDF search.
- **Progress Indicators**: Real-time updates on search and processing progress.
- **Multi-threading & Multiprocessing**: Efficient handling of file operations and PDF page counting to speed up the process.
- **Excel Reporting**: Save the results, including errors, to an Excel file for further analysis.

## Installation

1. **Clone the Repository**:
    ```bash
    git clone https://github.com/yourusername/pdf-page-counter.git
    ```

2. **Navigate to the Directory**:
    ```bash
    cd pdf-page-counter
    ```

3. **Install Dependencies**:
    Make sure you have Python installed. Install the required Python packages using:
    ```bash
    pip install -r requirements.txt
    ```

4. **Run the Application**:
    ```bash
    python pdf_page_counter.py
    ```

## Usage

1. **Open the Application**: Run the script to open the GUI.
2. **Select Folder**: Click on the "Select Folder" button to choose the directory containing your PDF files.
3. **Start Counting**: Click "Start Counting Pages" to begin the page counting process.
4. **View Results**: The results, including the total number of pages and PDFs, will be displayed in the application.
5. **Save to Excel**: The application will save a detailed report to an Excel file in the selected directory.

## Dependencies

- **Python 3.x**
- **Tkinter** (included with Python)
- **PyPDF2**
- **pandas**
- **openpyxl**

## Contributions

Feel free to fork this repository and submit pull requests. Any contributions to improve the application are welcome!

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Tags

- `Python`, `PDF`, `Tkinter`, `GUI`, `File Processing`, `Multi-threading`, `Multiprocessing`, `Automation`, `PDF Tools`, `PDF Page Counter`, `Desktop Application`, `Excel Report`, `PDF Analysis`
