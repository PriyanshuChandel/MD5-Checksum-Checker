# Generate Checksums for Files in a Directory
This is a Python GUI program that allows the user to select a directory using a file dialog and generate a checksum for all the files present in the selected directory and its subdirectories recursively. The program generates an Excel spreadsheet containing a list of file checksums.

### Requirements
Python 3.x
- `hashlib`: A Python library used for calculating hash values of files. Used to calculate the checksums for the files.
- `pathlib`: A Python library used for working with file system paths. Used to create and validate file paths.
- `tkinter`: Python's GUI package. Used to create the file selection dialog box and the GUI elements.
- `threading`: A Python library used to perform parallel execution of the program. Used to keep the GUI responsive while the program is executing.
- `time`: A Python library used for adding delays in the program. Used to add a delay to allow the GUI to update during program execution.
- `openpyxl`: A Python library used for working with Excel spreadsheets. Used to create and write data to the output Excel file.

To install these dependencies, you can use pip, the package installer for Python:
> `pip install hashlib pathlib tkinter openpyxl`

Note that threading, and time are built-in libraries and do not need to be installed separately.

### Installation and Usage
1. Clone the repository or download the zip file and extract it to your desired location.
2. Open the terminal/command prompt and navigate to the directory where the program is saved.
3. To run the program, enter the following command:
  > `python MD5_Checksum_Checker.py`
4. The GUI window will open. Click on the `...` button to select the directory for which you want to generate checksums.
5. Click on the `Submit` button to start the program. The program will start calculating the checksums for all the files in the selected directory and its subdirectories recursively.
6. Once the program finishes executing, it will generate an Excel spreadsheet containing a list of file checksums. The output Excel file will be saved in the same directory where the program is saved.

### Example
To generate checksums for all the files present in the "Documents" directory in your home folder, follow these steps:
1. Open the terminal/command prompt and navigate to the directory where the program is saved.
2. Enter the following command:
  > `python MD5_Checksum_Checker.py`
3. In the GUI window, click on the `...` button and select the "Documents" directory from your home folder.
4. Click on the `Submit` button to start the program.
5. Once the program finishes executing, an Excel spreadsheet containing a list of file checksums will be generated in the same directory where the program is saved.

### Contributions
Contributions to this repo are welcome. If you find a bug or have a suggestion for improvement, please open an issue on the repository. If you would like to make changes to the code, feel free to submit a pull request.

### Acknowledgments
This program was created as a part of a programming challenge. Special thanks to the challenge organizers for the inspiration.
