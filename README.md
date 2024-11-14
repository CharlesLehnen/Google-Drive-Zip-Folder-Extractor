# Google Drive Zip Folder Extractor

## Purpose

The University of Minnesota is deleting the Google Accounts of alumni, which will result in the loss of all data stored in alumni Google Drives. To make matters more challenging, Google Takeout and Google Transfer are disabled, forcing alumni to manually download Google Drive files as Zip folders. These Zip folders have size limits, leading to numerous separate files that make reconstructing the folder structure very time-consuming.

To assist with this, I've created this script, which can extract and rebuild the folder structure from the downloaded Zip files. The script is available as both a Python script and an executable file for Windows users. The executable was created to simplify the process for those without programming experience.

**Note:** This script has not been tested on Linux or macOS. If you try this script on a Linux or Mac, I would love to hear if it worked for you so I can update this note.

## Repository Structure

- **code/**: This directory contains the script (**code/extract_google_drive_output.ipynb**) to process the downloaded Google Drive files. 
- **dist/**: This directory contains the executable (`exe`) file for Windows users.

## Getting Started

### Requirements

- Python 3.8 or higher (if running from the script)

### Running the Project

There are three ways to run the project:

1. **Running Directly in Command Line**
   - Clone the repository.
   - Navigate to the code directory:
     ```sh
     python extract_google_drive_output.py
     ```

2. **Running via Executable File (Windows Only)**
   - Navigate to the `dist/` folder.
   - Locate the `.exe` file.
   - Run the `.exe` to start the program without needing to install Python or any dependencies.

3. **Running via Jupyter Notebook**

   - Clone the repository.
   - Navigate to the directory and open the Jupyter Notebook:
     ```sh
     jupyter notebook extract_google_drive_output.ipynb
     ```
   - Run the cells in the notebook to execute the extraction process. Alternatively you can just open in VSCode or other IDE and run there.

   Alternatively, you can run the `.ipynb` file using the command line:
   ```sh
   jupyter nbconvert --to notebook --execute extract_google_drive_output.ipynb
   ```

## Limitations

- The script is tested only on Windows, it has not been tested on Linux or macOS. If you try this script on a Linux or Mac, I would love to hear if it worked for you so I can update this note. I faced challenges because of the more restrictive path length and filename character limitations on Windows compared to Google Drive. This script addresses those issues by sanitizing paths and filenames as needed for compatibility with Windows. A CSV file is generated to detail the sanitization and renaming performed. I am unsure if these same requirements apply to Linux or Mac. So only if there's interest, I can work on making executables for Linux or macOS as well. If you try this script on a Linux or Mac, I would love to hear if it worked for you so I can update this note.

## Authors and Acknowledgment

- Charles Lehnen

## License

Refer to the `LICENSE` file for the full licensing agreement of the project.
