# **Folder Permissions Management Script for Windows**

This repository contains a simple yet effective solution for modifying the permissions of a folder and all its contents (files and subfolders) on Windows systems. It is particularly useful for resolving synchronization issues with cloud services like OneDrive, which sometimes fail due to restrictive permissions.

## **Why this Solution?**

This implementation, which separates the privilege elevation logic (VBScript) from the core script (Batch), offers several advantages:

* **Clarity:** Each file has a specific and well-defined role.  
* **Robustness:** Privilege elevation management via VBScript is reliable and does not require complex "hacks" within the batch script.  
* **Ease of Use:** Allows you to drag and drop a folder directly onto the executable to apply permissions, automatically launching as administrator.

## **Prerequisites**

* A Windows operating system (Windows 7 or later).  
* Access with an account that has permissions to perform administrative operations (UAC confirmation will be requested).

## **Solution Contents**

The solution consists of two distinct files that must be placed in the same directory:

1. launch\_as\_admin.vbs  
2. modify\_folder\_permissions.bat

### **1\. launch\_as\_admin.vbs**

This is a small VBScript that acts as a "wrapper" for the main batch script. Its sole purpose is to launch modify\_folder\_permissions.bat with elevated privileges (as administrator), ensuring that permission changes can be applied correctly.

### **2\. modify\_folder\_permissions.bat**

This is the main batch script that performs the permission modification operation. It accepts a folder path as an argument and uses the Windows icacls command to grant "full control" to the "Users" group on the specified folder and all its contents recursively.

## **How to Use**

1. **Download or Create Files:** Save the provided code for launch\_as\_admin.vbs and modify\_folder\_permissions.bat into two separate files with the indicated names.  
2. **Placement:** Ensure both files are in the **same folder** on your computer.  
3. **Execution (Recommended Method):**  
   * Drag the folder whose permissions you want to modify directly onto the launch\_as\_admin.vbs file icon.  
   * The system will prompt you for confirmation via **User Account Control (UAC)** to run the script as an administrator. Confirm to proceed.  
4. **Execution (Alternative Method):**  
   * Double-click the launch\_as\_admin.vbs file.  
   * If you haven't dragged a folder, the script will show you an error message and instructions on how to provide the folder path (e.g., by running it from the command line).

Once launched and UAC confirmed, the batch script will open, display the operation's status, and pause at the end, allowing you to read the messages before closing.

## **Technical Details (icacls)**

The script uses the icacls command with the following parameters:

* %TARGET\_FOLDER%: The path of the folder provided as input.  
* /grant Users:F: Grants "Full Control" (F) permission to the "Users" group. This allows all system users to have full control over the folder.  
* /T: Indicates to traverse all subdirectories and files within the specified folder, applying the modification recursively.  
* /C: Continues the operation even if errors occur (e.g., on locked or inaccessible files), attempting to apply permissions to as many elements as possible.  
* /Q: Suppresses success messages, making the output cleaner and focused on any errors.

## **Notes**

* Granting "full control" to the "Users" group is a rather broad action in terms of permissions. Ensure this is what you need for your specific situation (such as resolving synchronization issues). In more sensitive or multi-user environments, you might consider more granular permissions.  
* The script is designed to handle paths with spaces, thanks to proper argument handling in the VBScript and the use of quotes in the batch script.