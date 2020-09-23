VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Record Library"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9180
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   9180
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   7740
      TabIndex        =   56
      Top             =   6150
      Width           =   1245
   End
   Begin VB.Frame fraResults 
      Height          =   3285
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   9015
      Begin MSDataGridLib.DataGrid grdResults 
         Height          =   2955
         Left            =   120
         TabIndex        =   52
         Top             =   210
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   5212
         _Version        =   393216
         BorderStyle     =   0
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1043
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1043
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame TabFrame 
      BorderStyle     =   0  'None
      Height          =   2115
      Index           =   0
      Left            =   210
      TabIndex        =   14
      Top             =   3840
      Width           =   8805
      Begin VB.ComboBox cboFormats 
         Height          =   315
         ItemData        =   "frmMain.frx":08CA
         Left            =   3480
         List            =   "frmMain.frx":08CC
         TabIndex        =   55
         Top             =   1290
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.ComboBox cboYearOfRelease 
         Height          =   315
         Left            =   1560
         TabIndex        =   54
         Top             =   1290
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ComboBox cboLabel 
         Height          =   315
         Left            =   1560
         TabIndex        =   53
         Top             =   900
         Visible         =   0   'False
         Width           =   3285
      End
      Begin VB.TextBox txtFormat 
         Height          =   315
         Left            =   3480
         MaxLength       =   30
         TabIndex        =   5
         Top             =   1290
         Width           =   1365
      End
      Begin VB.TextBox txtArtist 
         Height          =   315
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   2
         Top             =   510
         Width           =   3285
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   375
         Left            =   2250
         TabIndex        =   8
         Top             =   1710
         Width           =   1245
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Edit"
         Height          =   375
         Left            =   3570
         TabIndex        =   12
         Top             =   1710
         Width           =   1245
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   7530
         TabIndex        =   11
         Top             =   1710
         Width           =   1245
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Height          =   375
         Left            =   6210
         TabIndex        =   10
         Top             =   1710
         Width           =   1245
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   375
         Left            =   4890
         TabIndex        =   9
         Top             =   1710
         Width           =   1245
      End
      Begin VB.TextBox txtTitle 
         Height          =   315
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   1
         Top             =   120
         Width           =   3285
      End
      Begin VB.TextBox txtLabel 
         Height          =   315
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   3
         Top             =   900
         Width           =   3285
      End
      Begin VB.TextBox txtYearOfRelease 
         Height          =   315
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   4
         Top             =   1290
         Width           =   1215
      End
      Begin VB.TextBox txtCatalogueNumber 
         Height          =   315
         Left            =   6360
         MaxLength       =   20
         TabIndex        =   6
         Top             =   120
         Width           =   2415
      End
      Begin VB.TextBox txtNotes 
         Height          =   1125
         Left            =   5520
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   480
         Width           =   3255
      End
      Begin VB.Label Label8 
         Caption         =   "Format:"
         Height          =   285
         Left            =   2880
         TabIndex        =   51
         Top             =   1320
         Width           =   645
      End
      Begin VB.Label lblTitle 
         Caption         =   "Title:"
         Height          =   285
         Left            =   330
         TabIndex        =   20
         Top             =   150
         Width           =   1245
      End
      Begin VB.Label Label2 
         Caption         =   "Notes:"
         Height          =   285
         Left            =   4890
         TabIndex        =   19
         Top             =   540
         Width           =   1245
      End
      Begin VB.Label Label3 
         Caption         =   "Catalogue Number:"
         Height          =   285
         Left            =   4890
         TabIndex        =   18
         Top             =   150
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Year of Release:"
         Height          =   285
         Left            =   330
         TabIndex        =   17
         Top             =   1320
         Width           =   1245
      End
      Begin VB.Label lblLabel 
         Caption         =   "Label:"
         Height          =   285
         Left            =   330
         TabIndex        =   16
         Top             =   930
         Width           =   1245
      End
      Begin VB.Label Label6 
         Caption         =   "Artist:"
         Height          =   285
         Left            =   330
         TabIndex        =   15
         Top             =   540
         Width           =   1245
      End
   End
   Begin VB.Frame TabFrame 
      BorderStyle     =   0  'None
      Height          =   2115
      Index           =   1
      Left            =   210
      TabIndex        =   21
      Top             =   3840
      Width           =   8805
      Begin VB.TextBox txtFormatSearch 
         Height          =   315
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   24
         Top             =   900
         Width           =   3285
      End
      Begin VB.CommandButton cmdCancelSearch 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   7530
         TabIndex        =   28
         Top             =   1710
         Width           =   1245
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         Height          =   375
         Left            =   6210
         TabIndex        =   26
         Top             =   1710
         Width           =   1245
      End
      Begin VB.TextBox txtTitleSearch 
         Height          =   315
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   22
         Top             =   120
         Width           =   3285
      End
      Begin VB.TextBox txtArtistSearch 
         Height          =   315
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   23
         Top             =   510
         Width           =   3285
      End
      Begin VB.TextBox txtLabelSearch 
         Height          =   315
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   25
         Top             =   1290
         Width           =   3285
      End
      Begin VB.Label lblFormatSearch 
         Caption         =   "Format:"
         Height          =   285
         Left            =   330
         TabIndex        =   58
         Top             =   930
         Width           =   1245
      End
      Begin VB.Label lblLabelSearch 
         Caption         =   "Label:"
         Height          =   285
         Left            =   330
         TabIndex        =   29
         Top             =   1320
         Width           =   1245
      End
      Begin VB.Label lblTitleSearch 
         Caption         =   "Title:"
         Height          =   285
         Left            =   330
         TabIndex        =   30
         Top             =   150
         Width           =   1245
      End
      Begin VB.Label lblArtistsSearch 
         Caption         =   "Artist:"
         Height          =   285
         Left            =   330
         TabIndex        =   27
         Top             =   540
         Width           =   1245
      End
   End
   Begin VB.Frame TabFrame 
      BorderStyle     =   0  'None
      Height          =   2115
      Index           =   3
      Left            =   210
      TabIndex        =   41
      Top             =   3840
      Width           =   8805
      Begin VB.ListBox lstLabels 
         Height          =   1035
         ItemData        =   "frmMain.frx":08CE
         Left            =   1560
         List            =   "frmMain.frx":08D0
         TabIndex        =   48
         Top             =   120
         Width           =   3285
      End
      Begin VB.TextBox txtNewLabel 
         Height          =   315
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   47
         Top             =   1230
         Width           =   3285
      End
      Begin VB.CommandButton cmdDeleteLabel 
         Caption         =   "Delete"
         Height          =   375
         Left            =   4890
         TabIndex        =   46
         Top             =   1710
         Width           =   1245
      End
      Begin VB.CommandButton cmdOKLabel 
         Caption         =   "OK"
         Height          =   375
         Left            =   6210
         TabIndex        =   45
         Top             =   1710
         Width           =   1245
      End
      Begin VB.CommandButton cmdCancelLabel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   7530
         TabIndex        =   44
         Top             =   1710
         Width           =   1245
      End
      Begin VB.CommandButton cmdEditLabel 
         Caption         =   "Edit"
         Height          =   375
         Left            =   3570
         TabIndex        =   43
         Top             =   1710
         Width           =   1245
      End
      Begin VB.CommandButton cmdAddLabel 
         Caption         =   "Add"
         Height          =   375
         Left            =   2250
         TabIndex        =   42
         Top             =   1710
         Width           =   1245
      End
      Begin VB.Label Label5 
         Caption         =   "Available Labels:"
         Height          =   555
         Left            =   330
         TabIndex        =   50
         Top             =   120
         Width           =   1155
      End
      Begin VB.Label Label7 
         Caption         =   "New Label:"
         Height          =   285
         Left            =   330
         TabIndex        =   49
         Top             =   1230
         Width           =   1245
      End
   End
   Begin VB.Frame TabFrame 
      BorderStyle     =   0  'None
      Height          =   2115
      Index           =   2
      Left            =   210
      TabIndex        =   31
      Top             =   3840
      Width           =   8805
      Begin VB.CommandButton cmdAddFormat 
         Caption         =   "Add"
         Height          =   375
         Left            =   2250
         TabIndex        =   38
         Top             =   1710
         Width           =   1245
      End
      Begin VB.CommandButton cmdEditFormat 
         Caption         =   "Edit"
         Height          =   375
         Left            =   3570
         TabIndex        =   37
         Top             =   1710
         Width           =   1245
      End
      Begin VB.CommandButton cmdCancelFormat 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   7530
         TabIndex        =   36
         Top             =   1710
         Width           =   1245
      End
      Begin VB.CommandButton cmdOKFormat 
         Caption         =   "OK"
         Height          =   375
         Left            =   6210
         TabIndex        =   35
         Top             =   1710
         Width           =   1245
      End
      Begin VB.CommandButton cmdDeleteFormat 
         Caption         =   "Delete"
         Height          =   375
         Left            =   4890
         TabIndex        =   34
         Top             =   1710
         Width           =   1245
      End
      Begin VB.TextBox txtNewFormat 
         Height          =   315
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   33
         Top             =   1230
         Width           =   3285
      End
      Begin VB.ListBox lstFormats 
         Height          =   1035
         ItemData        =   "frmMain.frx":08D2
         Left            =   1560
         List            =   "frmMain.frx":08D4
         TabIndex        =   32
         Top             =   120
         Width           =   3285
      End
      Begin VB.Label lblNewFormat 
         Caption         =   "New Format:"
         Height          =   285
         Left            =   330
         TabIndex        =   40
         Top             =   1230
         Width           =   1245
      End
      Begin VB.Label Label1 
         Caption         =   "Available Formats:"
         Height          =   555
         Left            =   330
         TabIndex        =   39
         Top             =   120
         Width           =   1155
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   2595
      Left            =   90
      TabIndex        =   13
      Top             =   3450
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   4577
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Details:"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Search:"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Formats:"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Labels:"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label lblAantal 
      Height          =   285
      Left            =   2130
      TabIndex        =   57
      Top             =   6180
      Width           =   4905
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Name: frmMain
' Author: Pascal Louwes
' Creation Date: 24-09-2001

Option Explicit

Private FillingGrid As Boolean
Private m_RecMode As RecordMode

Private Sub Form_Load()
    InitForm
    bEdit = False
    cmdAdd.Enabled = True
End Sub

Private Sub grdResults_Click()
    grdResults_RowColChange 0, 0
End Sub

Private Sub grdResults_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If FillingGrid Then
        Exit Sub
    Else
        Me.MousePointer = vbHourglass
        LoadFirstTab grdResults.Columns("RecordData_ID").Value
        DisableTextBoxes
        cmdAdd.Enabled = True
        cmdDelete.Enabled = True
        cmdEdit.Enabled = True
        cmdOK.Enabled = False
        cmdCancel.Enabled = False
        If cboFormats.Visible = True Then
            cboFormats.Visible = False
            txtFormat.Visible = True
            cboLabel.Visible = False
            txtLabel.Visible = True
            cboYearOfRelease.Visible = False
            txtYearOfRelease.Visible = True
        End If
        Me.MousePointer = vbNormal
    End If
End Sub
    
Private Sub Tabstrip1_Click()
    Dim iCurFrame As Integer ' Current Frame visible
    If TabStrip1.SelectedItem.Index = iCurFrame _
        Then Exit Sub
    TabFrame(0).Visible = False
    TabFrame(1).Visible = False
    TabFrame(2).Visible = False
    TabFrame(3).Visible = False
    TabFrame(TabStrip1.SelectedItem.Index - 1).Visible = True
    
    iCurFrame = TabStrip1.SelectedItem.Index
    
    Select Case iCurFrame
        Case 1
            'txtTitle.SetFocus
            cmdAdd.Default = True
        Case 2
            txtTitleSearch.SetFocus
            cmdSearch.Default = True
        Case 3
            lstFormats.SetFocus
            cmdAddFormat.Default = True
        Case 4
            lstLabels.SetFocus
            cmdAddLabel.Default = True
    End Select
    
End Sub
    
Private Sub cmdAdd_Click()
    PrepareTextBoxes
    cmdOK.Caption = "Save"
    cmdOK.Enabled = True
    cmdEdit.Enabled = False
    cmdDelete.Enabled = False
    cmdCancel.Enabled = True
End Sub
    
Private Sub cmdEdit_Click()
    EnableTextBoxes
    cmdAdd.Enabled = False
    cmdDelete.Enabled = False
    cmdEdit.Enabled = False
    cmdOK.Caption = "Save"
    cmdOK.Enabled = True
    cmdCancel.Enabled = True
    bEdit = True
    m_RecMode = rmEdit
End Sub
    
Private Sub cmdDelete_Click()
    Dim sResult As VbMsgBoxResult
    Dim sMsg As String
    
    On Error GoTo ErrorHandler
    
    sMsg = """" & txtTitle & """" & " by " & """" & txtArtist & """" & " is about to be deleted. Are you sure you want to continue?"
    sResult = MsgBox(sMsg, vbYesNo + vbExclamation, "Delete Record?")
    
    If sResult = vbYes Then
        Dim myRecords As clsRecords
        Set myRecords = New clsRecords
        
        myRecords.DeleteRecord grdResults.Columns("RecordData_ID").Value
        
        Set myRecords = Nothing
        InitForm
    Else
        Exit Sub
    End If
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbExclamation, "Error"
    Exit Sub
End Sub
    
Private Sub cmdOK_Click()
    If cmdOK.Caption = "Save" Then
        If txtTitle <> "" Then
            If txtArtist <> "" Then
                If Not cboLabel.Text = "choose an item" Then
                    If Not cboYearOfRelease.Text = "choose an item" Then
                        If Not cboFormats.Text = "choose an item" Then
                            
                            Dim myRecords As clsRecords
                            Set myRecords = New clsRecords
                            If bEdit = True Then
                                myRecords.SaveRecord txtTitle, _
                                    txtArtist, _
                                    cboFormats.ItemData(cboFormats.ListIndex), _
                                    cboLabel.ItemData(cboLabel.ListIndex), _
                                    cboYearOfRelease.ItemData(cboYearOfRelease.ListIndex), _
                                    txtCatalogueNumber, _
                                    txtNotes, _
                                    grdResults.Columns("RecordData_ID").Value
                            Else
                                myRecords.SaveRecord txtTitle, _
                                    txtArtist, _
                                    cboFormats.ItemData(cboFormats.ListIndex), _
                                    cboLabel.ItemData(cboLabel.ListIndex), _
                                    cboYearOfRelease.ItemData(cboYearOfRelease.ListIndex), _
                                    txtCatalogueNumber, _
                                    txtNotes
                            End If
                            cmdOK.Caption = "OK"
                            InitForm
                            Set myRecords = Nothing
                        Else
                            MsgBox "You cannot save this record because you did not choose a valid Format.", vbExclamation + vbOKOnly, "Choose valid Format"
                        End If
                    Else
                        MsgBox "You cannot save this record because you did not choose a valid Year of Release.", vbExclamation + vbOKOnly, "Choose valid Year"
                    End If
                Else
                    MsgBox "You cannot save this record because you did not choose a valid Label.", vbExclamation + vbOKOnly, "Choose valid Label"
                End If
            Else
                MsgBox "You cannot save this record because you did not fill in a valid Artist.", vbExclamation + vbOKOnly, "Enter Artist"
            End If
        Else
            MsgBox "You cannot save this record because you did not fill in a valid Title.", vbExclamation + vbOKOnly, "Enter Title"
        End If
    End If
End Sub
    
Private Sub cmdCancel_Click()
    InitForm
    With cmdAdd
        .Enabled = True
        .Default = True
    End With
End Sub

Private Sub cmdSearch_Click()
    
    If Not txtTitleSearch = "" Or Not txtArtistSearch = "" Or Not txtFormatSearch = "" Or Not txtLabelSearch = "" Then
        Dim myRecords As clsRecords
        Dim rsResultSet As ADODB.Recordset
        
        Set myRecords = New clsRecords
        Set rsResultSet = New ADODB.Recordset
        
        rsResultSet.Open myRecords.SearchRecord(txtTitleSearch, txtArtistSearch, txtFormatSearch, txtLabelSearch)
        
        FillingGrid = True
        Set grdResults.DataSource = rsResultSet
        grdResults.Columns("RecordData_ID").Visible = False
        FillingGrid = False
        
        Set rsResultSet = Nothing
        Set myRecords = Nothing
    Else
        MsgBox "Cannot perform search without parameters.", vbOKOnly + vbExclamation, "No search-parameters"
    End If
    
End Sub
    
Private Sub cmdCancelSearch_Click()
    InitForm
End Sub

Private Sub lstFormats_Click()
    cmdEditFormat.Enabled = True
    cmdDeleteFormat.Enabled = True
End Sub

Private Sub cmdAddFormat_Click()
    txtNewFormat.Enabled = True
    cmdOKFormat.Enabled = True
    cmdOKFormat.Caption = "Save"
    cmdCancelFormat.Enabled = True
End Sub
    
Private Sub cmdEditFormat_Click()
    bEdit = True
    txtNewFormat = lstFormats.Text
    txtNewFormat.Enabled = True
    txtNewFormat.SetFocus
    cmdOKFormat.Enabled = True
    cmdOKFormat.Caption = "Save"
    cmdDeleteFormat.Enabled = False
    cmdCancelFormat.Enabled = True
End Sub
    
Private Sub cmdOKFormat_Click()
    If txtNewFormat <> "" Then
        If cmdOKFormat.Caption = "Save" Then
            Dim myRecords As clsRecords
            Set myRecords = New clsRecords
            
            If bEdit = True Then
                myRecords.SaveAdditionalData "Formats", "Format", txtNewFormat, lstFormats.ItemData(lstFormats.ListIndex)
            Else
                myRecords.SaveAdditionalData "Formats", "Format", txtNewFormat
            End If
            cmdOK.Caption = "OK"
            txtNewFormat.Enabled = False
            txtNewFormat.Text = ""
            InitForm
            Set myRecords = Nothing
        End If
        bEdit = False
    Else
        MsgBox "Please fill in a valid Format", vbExclamation + vbOKOnly, "Fill in Format"
    End If
End Sub

Private Sub cmdDeleteFormat_Click()
    Dim sResult As VbMsgBoxResult
    Dim sMsg As String
    
    On Error GoTo ErrorHandler
    
    sMsg = "The format " & """" & lstFormats.Text & """" & " is about to be deleted. Are you sure you want to continue?"
    sResult = MsgBox(sMsg, vbOKCancel + vbExclamation, "Delete format?")
    
    If sResult = vbOK Then
        Dim myRecords As clsRecords
        Set myRecords = New clsRecords
        
        myRecords.DeleteAdditionalData "Formats", "Format_ID", lstFormats.ItemData(lstFormats.ListIndex)
        
        Set myRecords = Nothing
        InitForm
    Else
        Exit Sub
    End If
    cmdDeleteFormat.Enabled = False
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbExclamation, "Error"
    Exit Sub
End Sub
    
Private Sub cmdCancelFormat_Click()
    InitForm
End Sub
    
Private Sub lstLabels_Click()
    cmdEditLabel.Enabled = True
    cmdDeleteLabel.Enabled = True
End Sub

Private Sub cmdAddLabel_Click()
    txtNewLabel.Enabled = True
    cmdOKLabel.Enabled = True
    cmdOKLabel.Caption = "Save"
    cmdCancelLabel.Enabled = True
End Sub
    
Private Sub cmdEditLabel_Click()
    bEdit = True
    txtNewLabel = lstLabels.Text
    txtNewLabel.Enabled = True
    txtNewLabel.SetFocus
    cmdOKLabel.Enabled = True
    cmdOKLabel.Caption = "Save"
    cmdDeleteLabel.Enabled = False
    cmdCancelLabel.Enabled = True
End Sub

Private Sub cmdDeleteLabel_Click()
    Dim sResult As VbMsgBoxResult
    Dim sMsg As String
    On Error GoTo ErrorHandler
    
    sMsg = "The label " & """" & lstLabels.Text & """" & " is about to be deleted. Are you sure you want to continue?"
    sResult = MsgBox(sMsg, vbOKCancel + vbExclamation, "Delete label?")
    
    If sResult = vbOK Then
        Dim myRecords As clsRecords
        Set myRecords = New clsRecords
        
        myRecords.DeleteAdditionalData "Labels", "Label_ID", lstLabels.ItemData(lstLabels.ListIndex)
        
        Set myRecords = Nothing
        InitForm
    Else
        Exit Sub
    End If
    cmdDeleteLabel.Enabled = False
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbExclamation, "Error"
    Exit Sub
End Sub
    
Private Sub cmdOKLabel_Click()
    If txtNewLabel <> "" Then
        If cmdOKLabel.Caption = "Save" Then
            Dim myRecords As clsRecords
            Set myRecords = New clsRecords
            If bEdit = True Then
                myRecords.SaveAdditionalData "Labels", "Name", txtNewLabel, lstLabels.ItemData(lstLabels.ListIndex)
            Else
                myRecords.SaveAdditionalData "Labels", "Name", txtNewLabel
            End If
            cmdOK.Caption = "OK"
            txtNewLabel.Enabled = False
            txtNewLabel.Text = ""
            InitForm
            Set myRecords = Nothing
        End If
    Else
        MsgBox "Please fill in a valid Label", vbExclamation + vbOKOnly, "Fill in Label"
    End If
End Sub
    
Private Sub cmdCancelLabel_Click()
    InitForm
End Sub
    
Private Sub cmdClose_Click()
    Unload Me
    End
End Sub
    
Private Function InitForm()
    LoadGrid
    ClearTextBoxes
    FillListBox lstFormats
    FillListBox lstLabels
    bEdit = False
    cmdAddLabel.Enabled = True
    cmdAddFormat.Enabled = True
    txtNewLabel.Enabled = False
    txtNewFormat.Enabled = False
    cmdDelete.Enabled = False
    cmdDeleteLabel.Enabled = False
    cmdDeleteFormat.Enabled = False
    cmdEdit.Enabled = False
    cmdEditLabel.Enabled = False
    cmdEditFormat.Enabled = False
    cmdOK.Caption = "OK"
    cmdOK.Enabled = False
    cmdOKLabel.Enabled = False
    cmdOKFormat.Enabled = False
    cmdCancel.Enabled = False
    cmdCancelSearch.Enabled = False
    cmdCancelLabel.Enabled = False
    cmdCancelFormat.Enabled = False
End Function
    
Private Function DisableTextBoxes()
    txtTitle.Enabled = False
    txtArtist.Enabled = False
    txtLabel.Enabled = False
    txtYearOfRelease.Enabled = False
    txtFormat.Enabled = False
    txtCatalogueNumber.Enabled = False
    txtNotes.Enabled = False
End Function
    
Private Sub PrepareTextBoxes()
    With txtTitle
        .Enabled = True
        .Text = ""
    End With
    With txtArtist
        .Enabled = True
        .Text = ""
    End With
    With txtLabel
        .Enabled = True
        .Text = ""
        .Visible = False
        InitCombo cboLabel
    End With
    With txtYearOfRelease
        .Enabled = True
        .Text = ""
        .Visible = False
        InitCombo cboYearOfRelease
    End With
    With txtFormat
        .Enabled = True
        .Text = ""
        .Visible = False
        InitCombo cboFormats
    End With
    With txtCatalogueNumber
        .Enabled = True
        .Text = ""
    End With
    With txtNotes
        .Enabled = True
        .Text = ""
    End With
End Sub
    
Private Sub EnableTextBoxes()
    txtTitle.Enabled = True
    txtArtist.Enabled = True
    InitCombo cboLabel, grdResults.Columns("RecordData_ID").Value
    txtLabel.Visible = False
    InitCombo cboFormats, grdResults.Columns("RecordData_ID").Value
    txtFormat.Visible = False
    InitCombo cboYearOfRelease, grdResults.Columns("RecordData_ID").Value
    txtYearOfRelease.Visible = False
    txtCatalogueNumber.Enabled = True
    txtNotes.Enabled = True
End Sub
    
Private Function LoadGrid(Optional sTitle As String, _
    Optional sArtist As String, _
        Optional sLabel As String)
    
    If sTitle = "" And sArtist = "" And sLabel = "" Then
        Dim myRecords As clsRecords
        Set myRecords = New clsRecords
        Dim rsResultSet As ADODB.Recordset
        Set rsResultSet = New ADODB.Recordset
        
        rsResultSet.Open myRecords.GetRecordData(0)
        
        FillingGrid = True
        Set grdResults.DataSource = rsResultSet
        FillingGrid = False
        
        'lblAantal.Caption = "Er staan op dit moment " & rsResultSet.RecordCount & " titels in de database..."
        grdResults.Columns("RecordData_ID").Visible = False
        
        Set myRecords = Nothing
        Set rsResultSet = Nothing
    End If
End Function

Private Function LoadFirstTab(RecordIndex As Integer)
    Dim myRecords As clsRecords
    Set myRecords = New clsRecords
    Dim rsResultSet As ADODB.Recordset
    Set rsResultSet = New ADODB.Recordset
    
    rsResultSet.Open myRecords.GetRecordData(1, RecordIndex)
    
    txtTitle.Text = rsResultSet.Fields("Title").Value
    txtArtist.Text = rsResultSet.Fields("Artist").Value
    
    If Not IsNull(rsResultSet.Fields("Label").Value) Then
        txtLabel.Text = rsResultSet.Fields("Label").Value
    Else
        txtLabel.Text = "Unknown"
    End If
    
    If Not IsNull(rsResultSet.Fields("YearOfRelease").Value) Then
        txtYearOfRelease.Text = rsResultSet.Fields("YearOfRelease").Value
    Else
        txtYearOfRelease.Text = "Unknown"
    End If
    
    If Not IsNull(rsResultSet.Fields("Format").Value) Then
        txtFormat.Text = rsResultSet.Fields("Format").Value
    Else
        txtFormat.Text = "Unknown"
    End If
    
    If Not IsNull(rsResultSet.Fields("Catalogue_Number").Value) Then
        txtCatalogueNumber.Text = rsResultSet.Fields("Catalogue_Number").Value
    Else
        txtCatalogueNumber.Text = "Unknown"
    End If
    
    If Not IsNull(rsResultSet.Fields("Notes").Value) Then
        txtNotes.Text = rsResultSet.Fields("Notes").Value
    Else
        txtNotes.Text = "No notes available"
    End If
    
    Set myRecords = Nothing
    Set rsResultSet = Nothing
End Function
    
Private Function ClearTextBoxes()
    With txtTitle
        .Text = ""
        .Visible = True
    End With
    With txtArtist
        .Text = ""
        .Visible = True
    End With
    With txtLabel
        .Text = ""
        .Visible = True
    End With
    With txtYearOfRelease
        .Text = ""
        .Visible = True
    End With
    With txtFormat
        .Text = ""
        .Visible = True
    End With
    With txtCatalogueNumber
        .Text = ""
        .Visible = True
    End With
    With txtNotes
        .Text = ""
        .Visible = True
    End With
    DisableTextBoxes
    cboFormats.Visible = False
    cboLabel.Visible = False
    cboYearOfRelease.Visible = False
End Function

