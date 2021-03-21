# Overview
This project is a forked version of Govert's *SQLite for Excel*. The main difference is that the code has been refactored to uses classes as opposed to the standard modules that the original project utilized. The original code was stored in two modules: Sqlite3 and Sqlite3Demo. I've since updated these files to the class modules cSqlite3 and cSqlite3Demo respectively.

# Motivations for refactoring

1. There was a bug in the original code that was related to the directory the Excel file was stored in. If the current path was not the same path as the workbook's current path, the code would throw an error. To fix this, you'd have to add code to update the correct directory and call this method first or add code to each public method that does this. Code that does this must be run before any of those public methods can be utilized. This can be very easily done using a constructor with class modules, which is what I did. I've since added the private CheckOrChangeDir method which fixes this. So that was the primary motivation for refactoring to classes.
2. Using class modules allows you to use interfaces and polymorphism which make for easy refactoring, unit testing, etc. At some point in the future I will refactor the code to use interfaces so that this can be done.

# Known bugs

One of the test methods in cSqlite3Demo, TestBackup, throws an error. It is the only method among the original testing method that throws an error. I tested the method in the original code, and it also threw an error. So the error is not a result of refactoring. Since the method was throwing an error, the calling code in the AllTests method has been commented out. However, the original code for the method is still there.

# Potential bugs

This code may not be compatible with 32-bit versions of Office. If you need 32-bit compatability, I would recommend using the code in Govert's original project.

# Future updates

1. Future updates will include simpler and more concise code. As an example, the original code uses conditional compilation in various places to differentiate between the long and longPtr datatypes. However, [longPtr transforms to long in 32 bit, and longLong in 64 bit environments](https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/longptr-data-type). So the conditional compilation between one type to the other is not necessary.
2. Future versions will implement and utilize the ISqlite3 interface instead of making direct reference to the cSqlite3 class. Significant work on this update has been completed. However more work needs to be completed.
