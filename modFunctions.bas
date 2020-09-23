Attribute VB_Name = "modFunctions"
Option Explicit

Public Sub InitCombo(myCombo As ComboBox, Optional iSelectedRecordID As Integer)
    
    Dim myRecords As clsRecords
    Dim rsResultSet As ADODB.Recordset
    Dim sFieldName As String
    Dim sIDFieldName As String
    Dim iComboID As Integer
    
    Set myRecords = New clsRecords
    Set rsResultSet = New ADODB.Recordset
    With rsResultSet
        .Open myRecords.GetControlSettings(myCombo.Name, sFieldName, sIDFieldName)
        If Not .BOF And Not .EOF Then
            .MoveFirst
        Else
            'Exit Sub
        End If
    End With
    With myCombo
        .Clear
        If (rsResultSet.BOF And rsResultSet.EOF) Then
            .AddItem "no selection possible"
            .Enabled = False
        Else
            .AddItem "choose an item" 'Als je deze tekst in de combo hebben wil
        End If
        .ListIndex = 0 ' ...vergeet dan niet om hieraan een listindex en
        .ItemData(.NewIndex) = -1 '...een itemdata mee te geven....
        While Not rsResultSet.EOF '... en VB gaat vanaf hier zelf verder.
            .AddItem (rsResultSet.Fields(sFieldName).Value)
            .ItemData(.NewIndex) = rsResultSet.Fields(sIDFieldName).Value
            rsResultSet.MoveNext
        Wend
    End With
    
    myCombo.Visible = True
    
    Set rsResultSet = Nothing
    
    Set rsResultSet = New ADODB.Recordset
    
    If iSelectedRecordID <> 0 Then
        rsResultSet.Open myRecords.getComboIDs(iSelectedRecordID)
        setComboboxIndex myCombo, rsResultSet.Fields(sIDFieldName).Value, True
        Set rsResultSet = Nothing
    End If
    
End Sub

Public Sub FillListBox(myListbox As ListBox)
    Dim myRecords As clsRecords
    Dim rsResultSet As ADODB.Recordset
    Dim sFieldName As String
    Dim sIDFieldName As String
    
    Set myRecords = New clsRecords
    Set rsResultSet = New ADODB.Recordset
    
    With rsResultSet
        
        .Open myRecords.GetControlSettings(myListbox.Name, sFieldName, sIDFieldName)
        If Not .BOF And Not .EOF Then
            .MoveFirst
            myListbox.Clear
            While Not .EOF
                With myListbox
                    .AddItem rsResultSet.Fields(sFieldName).Value
                    .ItemData(.NewIndex) = rsResultSet.Fields(sIDFieldName).Value
                End With
                .MoveNext
            Wend
        End If
    End With
    
    Set myRecords = Nothing
    Set rsResultSet = Nothing
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Functie voor het "setten" van een combobox
'   Door Richard Scholten
'   06-09-01 'Sub van gemaakt, door Pascal.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub setComboboxIndex(ByRef ComboBox As ComboBox, ByVal Value As Variant, Optional ByVal UseItemdata As Boolean = True)
    Dim i As Integer
    Dim bFound As Boolean
    
    bFound = False
    For i = 0 To ComboBox.ListCount - 1
        If UseItemdata Then
            If ComboBox.ItemData(i) = Value Then
                bFound = True
                Exit For
            End If
        Else
            If ComboBox.List(i) = Value Then
                bFound = True
                Exit For
            End If
        End If
    Next i
    
    If bFound Then
        ComboBox.ListIndex = i
    Else
        ComboBox.ListIndex = 0
    End If
End Sub

