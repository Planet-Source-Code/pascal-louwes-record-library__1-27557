Attribute VB_Name = "modDBConnection"
Option Explicit
    
    Public cn_RecordLibrary As ADODB.Connection

Public Sub OpenDBConnectionRecordLibrary()
    If cn_RecordLibrary Is Nothing Then
        Set cn_RecordLibrary = New ADODB.Connection
        With cn_RecordLibrary
            .ConnectionString = DB_RecordLibraryConnectionString
            .CursorLocation = adUseClient
            .Open
        End With
    End If
End Sub

Public Sub CloseDBConnectionRecordLibrary()
    If Not cn_RecordLibrary Is Nothing Then
        If cn_RecordLibrary.State = adStateOpen Then
            cn_RecordLibrary.Close
        End If
        Set cn_RecordLibrary = Nothing
    End If
End Sub

