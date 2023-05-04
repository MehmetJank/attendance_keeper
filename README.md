# Attendance Keeper 

This is a simple Attendance Keeper App designed as a school project by Mehmet Can Kaya. The app is built using Python's tkinter module, openpyxl and xlwt libraries.

## Features

The Attendance Keeper App provides the following features:

- Import attendance data from an Excel file
- Select a section to view its students
- Mark students as attended or absent
- Export the list of attended students as a text or Excel file

## Screenshots

![attendance_keeper_1](https://user-images.githubusercontent.com/92443831/236318267-806f78d3-f8c5-43cd-88ca-7592ad9b9157.png)
![attendance_keeper_2](https://user-images.githubusercontent.com/92443831/236318278-7193e893-ab06-4cbd-9f4c-ec9b24efa55d.png)

## How to Use

    1. Open the app using the command python app.py
    2. Click on the "Import Excel File" button to import the attendance data from an Excel file. The file should have the following format:

      | ID | Name           | Department | Section |
      | -- | --------------| ----------| --------|
      | 1  | John Sebastian| CSE        | A       |
      | 2  | Jane Smith    | CSE        | A       |
      | 3  | Robert Johnson| EEE        | B       |

    3. After importing the data, select a section from the combobox.
    4. The list of students in the selected section will be displayed in the left list box.
    5. To mark a student as attended, select the student's name from the list and click on the "Add" button.
    6. The list of attended students can be exported as a text or Excel file by clicking on the "Export" button and selecting the desired file format.

## Requirements

The Attendance Keeper App requires the following libraries to be installed:

- tkinter
- openpyxl
- xlwt

To install the libraries, run the following commands:

```
pip install tkinter openpyxl xlwt
```

## Credits

The Attendance Keeper App was developed by Mehmet Can Kaya for school project purposes. 
