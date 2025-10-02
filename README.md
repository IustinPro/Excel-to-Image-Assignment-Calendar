# Excel-to-Image-Assignment-Calendar
A simple program that takes an excel table and turns it into an easy to read image.

Steps of Use:

1.
In folder dedicated to program you must have as follows:
excel_to_image.py
assignments.xlsx

The Assignments Excel must be formatted as follows (or use provided Template):
Unit Nr,	Assignment Name,	Assignment Number,	Due in

Unit Nr & Assignment Number can only be a number.
Due in Date must be formatted as DD/MM/YYYY e.g 12 January 2025 would be 12/01/2025, Only Numerical values
Additionaly there is support for the due in date to be TBD (To Be Decided)

2.
Copy your address to the folder and in powershell run

   cd C:\Users\given_user\path\path\folder

3. then to run the script in powershell run

   python excel_to_image.py

This should then output Saved image to assignments.png (roughly width=816, height=310)

4. Refresh the folder and now there is a png formatted for you!

5. To reuse don't forget to delete the existing png.
