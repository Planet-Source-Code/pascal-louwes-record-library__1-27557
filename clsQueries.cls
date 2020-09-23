VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsQueries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function GetQuery(sFunction As String) As String
    
    Dim iTableID As Integer
    Dim iSecondTableID As Integer
    Dim sSql As String
    Dim sSqlTwo As String
    Dim bTwoParts As Boolean
    Dim sReturn As String
    
    Select Case sFunction
        Case "GetRecordData"
            iTableID = 1
            iSecondTableID = 2
            bTwoParts = True
        Case "SaveRecord"
            iTableID = 3
        Case "DeleteRecord"
            iTableID = 4
        Case "SearchRecord"
            iTableID = 5
            iSecondTableID = 6
            bTwoParts = True
        Case "SaveDeleteFormats"
            iTableID = 7
        Case "SaveDeleteLabels"
            iTableID = 8
        Case "GetControlLabel"
            iTableID = 9
        Case "GetControlFormat"
            iTableID = 10
        Case "GetControlYOR"
            iTableID = 11
        Case "getComboIDs"
            iTableID = 12
            iSecondTableID = 13
            bTwoParts = True
    End Select
    
    Dim rsQuery As ADODB.Recordset
    
    OpenDBConnectionRecordLibrary
    
    Set rsQuery = New ADODB.Recordset
    
    sSql = "Select SqlQuery From SqlQueries Where SqlQueryID = " & iTableID
    
    rsQuery.Open sSql, cn_RecordLibrary, adOpenForwardOnly, adLockReadOnly
    
    sReturn = rsQuery.Fields("SqlQuery").Value
    
    rsQuery.Close
    
    If bTwoParts Then 'If query is longer than 255 chs, it's devided over two rows in the DB
        sSqlTwo = "Select SqlQuery From SqlQueries Where SqlQueryID = " & iSecondTableID
        rsQuery.Open sSqlTwo, cn_RecordLibrary, adOpenForwardOnly, adLockReadOnly
        sReturn = sReturn & " " & rsQuery.Fields("SqlQuery").Value
        bTwoParts = False
    End If
    
    Set rsQuery.ActiveConnection = Nothing
    
    CloseDBConnectionRecordLibrary
    
    GetQuery = sReturn
    
End Function

