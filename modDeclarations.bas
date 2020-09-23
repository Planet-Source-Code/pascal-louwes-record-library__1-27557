Attribute VB_Name = "modDeclarations"
Option Explicit

' Name:
' Author: Pascal Louwes
' Creation Date: 24-09-2001

Public Const DB_RecordLibraryConnectionString As String = "DSN=RecordLibrary;UID=;PWD="
Public bEdit As Boolean

Public Enum SearchType
    stAllRecords = 0
    stSelectedRecord = 1
End Enum

Public Enum RecordMode
    rmAdd = 0
    rmEdit = 1
    rmView = 2
    rmDelete = 3
End Enum

