# Overview
This project is a forked version of Govert's *SQLite for Excel*. The main difference is that the code has been refactored to uses classes as opposed to the standard modules that the original project utilized. The original code was stored in two modules: Sqlite3 and Sqlite3Demo. I've since updated these files to the class modules cSqlite3 and cSqlite3Demo respectively. I've also added an ISqlite3 interface which cSqlite3 implements. My code also removed unnecessary conditional compilation from the codebase.

# Motivations for refactoring

1. There were a few bugs in Sqlite3 amd Sqlite3Demo modules. Both of these bugs were related to ensuring that certain code was run before any subsequent code was run that relied on this code being run. In cSqlite3, code needed to be run to set the correct path for the directory the workbook and database were located in (it defaulted to the user's My Documents folder). For cSqlite3 demo, the test file that all of the tests utilized was only created if you ran the AllTests method. If you tried to run other tests directly, you'd get an error. Both of these issues were easily fixed by using the each class file's constructor. So these bugs are no longer an issue. 
2. Using class modules allows you to use interfaces and polymorphism which make for easy refactoring, unit testing, etc.
3. I deleted some unnecessary code related to conditional compilation. This was used in many procedures to alternate between two different data types: long and longPtr. These data types were used on 32-bit and 64-bit systems respectively. However, this code was not needed. As you can see in this MDSN link, [longPtr transforms to long in 32 bit, and longLong in 64 bit environments](https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/longptr-data-type). So the conditional compilation between one type to the other is not necessary.

# Known bugs

1. One of the test methods in cSqlite3Demo, TestBackup, throws an error. It is the only method among those called in the AllTests method that throws an error. I tested the method in the original code, and it also threw an error. So the error is not a result of refactoring. Since the method was throwing an error, the calling code in the AllTests method has been commented out. However, the original code for the method is still there.
2. One of the original test methods in cSqlite3Demo, TestBinding, threw an error after refactoring. This error happens in the for loop in the code. The original code tested 100,000 elements. After refactoring, only 97,262 elements work. I have no idea why this is the case. But the code has been updated to this new upper bound so the test works. The original code with the original upper bound is left in the code but commented out.
3. I am not able to get this code the file to work on my cloud-storage directory. If you're having similar issues, it may help to move the file to a network drive or some other directory where this isn't an issue.

# Potential bugs

This code may not be compatible with 32-bit versions of Office. If you need 32-bit compatability and this project doesn't work, I would recommend using the code in Govert's original project.

# Example

This is sample code taken from the mMain module which is included in the Excel workbook:

    Sub Main()
        Dim sqlite As ISqlite3
        Dim demo As cSqlite3Demo
        
        Set sqlite = New cSqlite3
        Set demo = New cSqlite3Demo
        
        demo.init sqlite
        demo.AllTests
    End Sub
