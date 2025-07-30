# **PowerPoint to PDF Batch Converter**

This application provides a convenient way to convert multiple PowerPoint presentation files (.pptx, .ppt) within a specified folder to PDF format. It automates the conversion process, making it ideal for bulk conversions and allowing for multiple conversion sessions without restarting the script.

## **Why this Solution?**

This solution uses a PowerShell script for the core conversion logic, orchestrated by a simple batch file. This approach offers:

* **Automation:** Converts multiple files with a single command.  
* **Simplicity:** Easy to use by simply running a batch file and providing a folder path.  
* **Organization:** Automatically creates a dedicated PDF subfolder within your source directory for the converted files.  
* **Error Handling:** Includes basic checks for prerequisites and conversion issues.  
* **Batch Processing Loop:** Allows you to convert files from multiple folders in a single session without restarting the application.

## **Prerequisites**

* **Microsoft PowerPoint:** This script requires a full installation of Microsoft PowerPoint (2010 or later is recommended) on the system where the script is executed. The script interacts with PowerPoint's COM object model to perform the conversions.  
* **Windows Operating System:** Designed for Windows environments.  
* **PowerShell 5.1 or later:** The Convert-PPTXtoPDF.ps1 script requires PowerShell version 5.1 or newer. This is typically pre-installed on modern Windows systems.

## **Solution Contents**

The application consists of two files that must be placed in the same directory:

1. Convert.bat  
2. Convert-PPTXtoPDF.ps1

### **1\. Convert.bat**

This is the main batch file that you will execute. It serves as a user-friendly interface:

* It prompts the user to enter the path to the folder containing the PowerPoint files.  
* It performs basic validation to ensure the PowerShell script exists and the provided folder path is valid.  
* It then calls the Convert-PPTXtoPDF.ps1 script, passing the specified source folder as an argument.  
* **New Feature:** After each conversion session, it will ask the user if they wish to convert more files, allowing for continuous operation.

### **2\. Convert-PPTXtoPDF.ps1**

This is the core PowerShell script responsible for the conversion process:

* It takes the SourceFolder path as a mandatory parameter.  
* It checks if Microsoft PowerPoint is installed and can be accessed.  
* It creates a PDF subfolder within the SourceFolder if it doesn't already exist.  
* It iterates through all .pptx and .ppt files found directly in the SourceFolder.  
* For each file, it opens the presentation in PowerPoint (in read-only mode) and saves it as a PDF.  
* It skips files that have already been converted (if a PDF with the same name exists).  
* It includes error handling for presentations that might be open, password-protected, or corrupted.  
* Finally, it quits the PowerPoint application and releases COM objects to ensure proper termination.

## **How to Use**

1. **Download or Create Files:** Save the code for Convert.bat and Convert-PPTXtoPDF.ps1 into two separate files with the indicated names.  
2. **Placement:** Ensure both files are in the **same folder** on your computer.  
3. **Execute the Batch File:**  
   * Double-click the Convert.bat file.  
   * A command prompt window will open, guiding you through the process.  
   * You will be prompted to "Please enter the full path to the folder containing your PowerPoint files:".  
   * Enter the full path (e.g., C:\\Users\\YourUser\\Documents\\Presentations) and press Enter.  
4. **Conversion Process:**  
   * The script will then start the conversion process. You will see messages indicating which files are being converted, skipped, or if any errors occur.  
   * Converted PDF files will be saved in a new PDF subfolder within the source folder you specified.  
5. **Continue or Exit:**  
   * After the conversion of the current folder is finished, the script will ask: Do you want to convert more files? (Y/N):  
   * Type Y (or y) and press Enter to start another conversion session for a different folder.  
   * Type N (or n) and press Enter to exit the converter.  
   * If you enter an invalid choice, it will prompt you again.

## **Troubleshooting**

* **"PowerPoint is not installed or cannot be started":** Ensure Microsoft PowerPoint is correctly installed on your system.  
* **"Failed to convert... It might be open, password-protected, or corrupted":**  
  * Close any open PowerPoint files before running the script.  
  * The script cannot convert password-protected presentations.  
  * Corrupted files may also fail to convert. Try opening them manually in PowerPoint to check for issues.  
* **"Error: The source folder was not found":** Double-check that the path you entered is correct and the folder exists.  
* **No PDFs are created:**  
  * Verify that .pptx or .ppt files exist directly within the specified SourceFolder.  
  * Check the command prompt output for any error messages.

## **Notes**

* The script processes files non-recursively within the SourceFolder itself, meaning it will not convert files in subfolders *within* the SourceFolder.  
* The script automatically handles the creation of the PDF output directory.  
* The script attempts to clean up PowerPoint processes after conversion, but in rare cases, you might need to manually close PowerPoint if it remains open.