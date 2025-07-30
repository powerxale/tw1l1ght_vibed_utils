# **tw1l1ght\_vibed\_utils: A Collection of Practical Windows Utilities**

This repository houses a collection of small, specific utilities designed to streamline common tasks on Windows operating systems. These tools are developed with a focus on simplicity, automation, and direct usability.

## **Crafted with "Vibe Coding"**

Each utility within this collection has been created through an iterative and collaborative "vibe coding" process, where the user's specific needs and desired functionalities were translated into practical scripts. As an AI, Google Gemini had the privilege of assisting in the design, development, and refinement of these tools, ensuring they meet the my requirements for efficiency and ease of use.

## **Utilities Included**

### **1\. PowerPoint to PDF Batch Converter**

This application provides a convenient way to convert multiple PowerPoint presentation files (.pptx, .ppt) within a specified folder to PDF format. It automates the conversion process, making it ideal for bulk conversions and allowing for multiple conversion sessions without restarting the script.

* **Purpose:** Convert multiple PowerPoint files to PDF efficiently.  
* **Key Features:**  
  * Batch conversion with a single command.  
  * Automatic creation of a PDF subfolder for converted files.  
  * Error handling for common issues (e.g., missing PowerPoint, corrupted files).  
  * Interactive loop to convert files from multiple folders in one session.  
* **Prerequisites:** Microsoft PowerPoint (2010 or later), Windows OS, PowerShell 5.1 or later.  
* **Files:** Convert.bat, Convert-PPTXtoPDF.ps1  
* **Usage:** Double-click Convert.bat and follow the prompts, or refer to the dedicated PPTX-to-PDF Converter/README.md for detailed instructions.

### **2\. Folder Permissions Management Script**

This utility provides a simple yet effective solution for modifying the permissions of a folder and all its contents (files and subfolders) on Windows systems. It is particularly useful for resolving synchronization issues with cloud services like OneDrive, which sometimes fail due to restrictive permissions.

* **Purpose:** Grant "Full Control" permissions to the "Users" group on a specified folder and its contents.  
* **Key Features:**  
  * Elevates privileges automatically (requires UAC confirmation).  
  * Recursively applies permissions to all files and subfolders.  
  * Designed for drag-and-drop usage for ease.  
* **Prerequisites:** Windows OS, administrative privileges (UAC will prompt).  
* **Files:** launch\_as\_admin.vbs, modify\_folder\_permissions.bat  
* **Usage:** Drag and drop the target folder onto launch\_as\_admin.vbs, or refer to the dedicated Change folder rights/README.md for detailed instructions.

## **General Notes**

* All utilities are designed for **Windows operating systems**.  
* Some utilities require **administrative privileges** to function correctly, and will prompt for User Account Control (UAC) confirmation.  
* For detailed usage instructions, troubleshooting, and specific prerequisites for each utility, please refer to their respective README.md files located in their subdirectories.

We hope these utilities prove useful in streamlining your daily tasks\!