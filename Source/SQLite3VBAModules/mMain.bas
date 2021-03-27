Attribute VB_Name = "mMain"
Option Explicit

Sub Main()
    Dim sqlite As ISqlite3
    Dim demo As cSqlite3Demo
    
    Set sqlite = New cSqlite3
    Set demo = New cSqlite3Demo
    
    demo.init sqlite
    demo.AllTests
End Sub
