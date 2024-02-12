# excel_to_csv

Excel to CSV Converter

  This Python script utilizes PyQt5 to create a user-friendly GUI application that converts Excel files (.xlsx) to CSV format. The application allows users to either drag and drop Excel files into the interface or 
  use a file dialog to select and convert files. It's designed to be simple and efficient, providing immediate feedback on the conversion status.

Features:
  Drag and drop functionality for Excel files.
  File dialog for selecting Excel files manually.
  Instant conversion of Excel (.xlsx) files to CSV.
  User feedback on successful conversions or errors.
  A clean, minimalistic interface for easy use.
  
Prerequisites:

  Before you begin, ensure you have the following requirements:

    Python 3.x installed
    PyQt5 installed
    Pandas library installed

  You can install the required libraries using pip:
    
    pip install PyQt5 pandas

Setup:
  Clone the repository or download the script to your local machine.
  Ensure all required libraries are installed.
  
Running the Application:

  To run the script, navigate to the script's directory in your terminal and execute:

    python excel_to_csv_converter.py
    

Usage:

  Upon running the application, you will see a window titled 'CONVERSOR DE EXCEL A CSV'. You can either click the button to open a file dialog and select an Excel file or drag and drop an Excel file directly onto the button. The application will immediately convert the file to CSV format and save it in the same directory as the original Excel file. The window will display the path to the converted CSV file or an error message if the conversion fails.

Code Explanation:
  
  PyQt5 for GUI: The script uses PyQt5 to build the graphical user interface, making it easy for users to interact with the application.
  Drag and Drop: Implements drag and drop functionality for a user-friendly experience.
  Pandas for Conversion: Utilizes Pandas for the conversion process, ensuring accurate and efficient conversion from Excel to CSV.
  File Dialog: Includes a file dialog for selecting files manually, offering flexibility in file selection.
  Error Handling: Provides feedback to the user in case of conversion errors, enhancing the user experience.
