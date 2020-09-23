VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRecords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Name: GetRecordData
' Author: Pascal Louwes
' Creation Date: 24-09-2001

Public Function GetRecordData(SearchType As Integer, Optional iRecordDataID As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim rsRecordData As ADODB.Recordset
    Dim Query As clsQueries
    Set Query = New clsQueries
    
    sSql = Query.GetQuery("GetRecordData")
    
    Set Query = Nothing
    
    If SearchType = 1 Then 'Client wants specific record. (Otherwise all records are returned)
        sSql = sSql & " AND RD.RecordData_ID = " & iRecordDataID
    End If
    
    OpenDBConnectionRecordLibrary
    
    Set rsRecordData = New ADODB.Recordset
    
    rsRecordData.Open sSql, cn_RecordLibrary, adOpenForwardOnly, adLockReadOnly
    
    Set rsRecordData.ActiveConnection = Nothing
    
    CloseDBConnectionRecordLibrary
    
    Set GetRecordData = rsRecordData
    
    Set rsRecordData = Nothing
    
End Function

' Name: SaveRecord
' Author: Pascal Louwes
' Creation Date: 24-09-2001

Public Sub SaveRecord(Optional sTitle As String, _
    Optional sArtist As String, _
        Optional iFormatID As Integer, _
        Optional iLabelID As Integer, _
        Optional iYearOfReleaseID As Integer, _
        Optional sCatalogueNumber As String, _
        Optional sNotes As String, _
        Optional iRecordDataID As Integer)
    
    Dim rsRecordData As ADODB.Recordset
    
    Set rsRecordData = New ADODB.Recordset
    
    If bEdit = True Then
        Dim sSql As String
        Dim Query As clsQueries
        
        Set Query = New clsQueries
        sSql = Query.GetQuery("SaveRecord") & iRecordDataID
        Set Query = Nothing
        
        OpenDBConnectionRecordLibrary
        
        rsRecordData.Open sSql, cn_RecordLibrary, adOpenStatic, adLockOptimistic
    Else
        OpenDBConnectionRecordLibrary
        
        With rsRecordData
            .Open "RecordData", cn_RecordLibrary, adOpenStatic, adLockOptimistic, adCmdTable
            .AddNew
        End With
    End If
    With rsRecordData
        .Fields("Title").Value = sTitle
        .Fields("Artist").Value = sArtist
        .Fields("Catalogue_Number").Value = IIf(sCatalogueNumber = "", Null, sCatalogueNumber)
        .Fields("Notes").Value = IIf(sNotes = "", Null, sNotes)
        .Fields("Format_ID").Value = IIf(iFormatID = Null, Null, iFormatID)
        .Fields("Label_ID").Value = IIf(iLabelID = Null, Null, iLabelID)
        .Fields("YearOfRelease_ID").Value = IIf(iYearOfReleaseID = Null, Null, iYearOfReleaseID)
        .Update
        .Close
    End With
    Set rsRecordData.ActiveConnection = Nothing
    
    CloseDBConnectionRecordLibrary
    
    Set rsRecordData = Nothing
    
End Sub

' Name: DeleteRecord
' Author: Pascal Louwes
' Creation Date: 24-09-2001

Public Sub DeleteRecord(iRecordDataID As Integer)
    Dim rsRecordData As ADODB.Recordset
    Dim sSql As String
    Set rsRecordData = New ADODB.Recordset
    
    Dim Query As clsQueries
    Set Query = New clsQueries
    
    sSql = Query.GetQuery("DeleteRecord") & iRecordDataID
    
    Set Query = Nothing

    OpenDBConnectionRecordLibrary
    
    With rsRecordData
        .Open sSql, cn_RecordLibrary, adOpenStatic, adLockOptimistic
        .Delete
        .Close
    End With
    Set rsRecordData.ActiveConnection = Nothing
    
    CloseDBConnectionRecordLibrary
    
    Set rsRecordData = Nothing
    
End Sub

' Name: SearchRecord
' Author: Pascal Louwes
' Creation Date: 24-09-2001

Public Function SearchRecord(Optional sTitle As String, _
    Optional sArtist As String, _
        Optional sFormat As String, _
        Optional sLabel As String) As ADODB.Recordset
    
    Dim sSql As String
    Dim sSqlAnd As String
    Dim rsRecordData As ADODB.Recordset
    
    Dim Query As clsQueries
    Set Query = New clsQueries
    
    sSql = Query.GetQuery("SearchRecord")
    
    Set Query = Nothing
    
    If Not sTitle = "" Then
        sSqlAnd = " AND RD.Title LIKE '%" & sTitle & "%' "
    End If
    
    If Not sArtist = "" Then
        sSqlAnd = sSqlAnd & " AND RD.Artist LIKE '%" & sArtist & "%' "
    End If
    
    If Not sLabel = "" Then
        sSqlAnd = sSqlAnd & " AND Labels.Name LIKE '%" & sLabel & "%' "
    End If
    
    If Not sFormat = "" Then
        sSqlAnd = sSqlAnd & " AND Formats.Format LIKE '%" & sFormat & "%' "
    End If
    
    sSql = sSql & sSqlAnd & "ORDER BY RD.Title"
    
    Set rsRecordData = New ADODB.Recordset
    
    OpenDBConnectionRecordLibrary
    
    rsRecordData.Open sSql, cn_RecordLibrary, adOpenForwardOnly, adLockReadOnly
    
    Set rsRecordData.ActiveConnection = Nothing
    
    CloseDBConnectionRecordLibrary
    
    Set SearchRecord = rsRecordData
    
    Set rsRecordData = Nothing
    
End Function

' Name: SaveAdditionalData
' Author: Pascal Louwes
' Creation Date: 24-09-2001

Public Sub SaveAdditionalData(sTableName As String, sFieldType As String, sData As String, Optional iRecordDataID As Integer)
    Dim rsRecordData As ADODB.Recordset
    
    Set rsRecordData = New ADODB.Recordset
    
    If bEdit = True Then
        Dim sSql As String
        Dim Query As clsQueries
        Set Query = New clsQueries
        
        sSql = Query.GetQuery("SaveDelete" & sTableName) & iRecordDataID
         
        Set Query = Nothing
        
        OpenDBConnectionRecordLibrary
        
        rsRecordData.Open sSql, cn_RecordLibrary, adOpenStatic, adLockOptimistic
        
    Else
        OpenDBConnectionRecordLibrary
        With rsRecordData
            .Open sTableName, cn_RecordLibrary, adOpenStatic, adLockOptimistic, adCmdTable
            .AddNew
        End With
    End If
    
    With rsRecordData
        .Fields(sFieldType).Value = sData
        .Update
        .Close
    End With
    
    Set rsRecordData.ActiveConnection = Nothing
    
    CloseDBConnectionRecordLibrary
    
    Set rsRecordData = Nothing
End Sub

' Name: DeleteAdditionalData
' Author: Pascal Louwes
' Creation Date: 24-09-2001

Public Sub DeleteAdditionalData(sTableName As String, sIDFieldName As String, iRecordDataID As Integer)
    Dim rsRecordData As ADODB.Recordset
    Dim sSql As String
    
    Dim Query As clsQueries
    Set Query = New clsQueries
    
    sSql = Query.GetQuery("SaveDelete" & sTableName) & iRecordDataID
    
    Set Query = Nothing
    
    OpenDBConnectionRecordLibrary
    
    Set rsRecordData = New ADODB.Recordset
    
    With rsRecordData
        .Open sSql, cn_RecordLibrary, adOpenStatic, adLockOptimistic
        .Delete
        .Close
    End With
    Set rsRecordData.ActiveConnection = Nothing
    
    CloseDBConnectionRecordLibrary
    
    Set rsRecordData = Nothing
    
End Sub

' Name:
' Author: Pascal Louwes
' Creation Date: 24-09-2001

Public Function GetControlSettings(sName As String, ByRef sDisplayField As String, Optional ByRef sDisplayID As String) As ADODB.Recordset
    Dim sSql As String
    Dim rsRecordData As ADODB.Recordset
    
    Dim Query As clsQueries
    Set Query = New clsQueries
    
    Select Case sName
        Case "cboLabel", "lstLabels": sSql = Query.GetQuery("GetControlLabel"): sDisplayField = "Name": sDisplayID = "Label_ID"
        Case "cboFormats", "lstFormats": sSql = Query.GetQuery("GetControlFormat"): sDisplayField = "Format": sDisplayID = "Format_ID"
        Case "cboYearOfRelease": sSql = Query.GetQuery("GetControlYOR"): sDisplayField = "YearOfRelease": sDisplayID = "YearOfRelease_ID"
    End Select
    
    Set Query = Nothing
    
    OpenDBConnectionRecordLibrary
    
    Set rsRecordData = New ADODB.Recordset
    
    rsRecordData.Open sSql, cn_RecordLibrary, adOpenForwardOnly, adLockReadOnly
    
    Set rsRecordData.ActiveConnection = Nothing
    
    CloseDBConnectionRecordLibrary
    
    Set GetControlSettings = rsRecordData
    
    Set rsRecordData = Nothing
    
End Function

' Name: getComboIDs
' Author: Pascal Louwes
' Creation Date: 24-09-2001

Public Function getComboIDs(ByVal RecordDataID As Integer) As ADODB.Recordset
    Dim sSql As String
    Dim rsRecordData As ADODB.Recordset
    Dim Query As clsQueries
    Set Query = New clsQueries
    
    sSql = Query.GetQuery("getComboIDs") & RecordDataID
    
    Set Query = Nothing
    
    OpenDBConnectionRecordLibrary
    
    Set rsRecordData = New ADODB.Recordset
    
    rsRecordData.Open sSql, cn_RecordLibrary, adOpenForwardOnly, adLockReadOnly
    
    Set rsRecordData.ActiveConnection = Nothing
    
    CloseDBConnectionRecordLibrary
    
    Set getComboIDs = rsRecordData
    
    Set rsRecordData = Nothing
    
End Function

