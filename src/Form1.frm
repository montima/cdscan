VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "CDScan"
   ClientHeight    =   5685
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8685
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5685
   ScaleWidth      =   8685
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List5 
      Height          =   840
      Left            =   1800
      TabIndex        =   19
      Top             =   3240
      Visible         =   0   'False
      Width           =   5655
   End
   Begin VB.ListBox List7 
      Height          =   1230
      Left            =   4920
      TabIndex        =   22
      Top             =   2520
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.ListBox List6 
      Height          =   1230
      Left            =   2520
      TabIndex        =   21
      Top             =   2520
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   495
      Left            =   7320
      TabIndex        =   20
      Top             =   4320
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   495
      Left            =   0
      TabIndex        =   18
      Top             =   840
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   873
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   16
      Top             =   5430
      Width           =   8685
      _ExtentX        =   15319
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.ToolTipText     =   "File Count"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton command2 
      Caption         =   "Scan New CD"
      Height          =   495
      Left            =   7320
      TabIndex        =   14
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   7320
      TabIndex        =   11
      ToolTipText     =   "Enter a Unique Disk Name"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton command4 
      Caption         =   "Add to Catalog..."
      Height          =   495
      Left            =   7320
      TabIndex        =   10
      Top             =   2160
      Width           =   1215
   End
   Begin VB.ListBox List2 
      Height          =   1425
      ItemData        =   "Form1.frx":1CFA
      Left            =   2520
      List            =   "Form1.frx":1D01
      TabIndex        =   8
      Top             =   960
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   2760
      TabIndex        =   7
      Top             =   4920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.DirListBox Dir1 
      Height          =   315
      Left            =   1440
      TabIndex        =   5
      Top             =   4920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton command3 
      Caption         =   "Write Catalog..."
      Height          =   495
      Left            =   7320
      TabIndex        =   4
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   1260
      Left            =   6120
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   5880
      TabIndex        =   0
      Top             =   4080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   720
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1D0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":25E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2EC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":37A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":407C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4958
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5234
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.ListBox List1 
      Height          =   1425
      ItemData        =   "Form1.frx":5B10
      Left            =   120
      List            =   "Form1.frx":5B17
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   4080
      TabIndex        =   6
      Top             =   4920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   810
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8685
      _ExtentX        =   15319
      _ExtentY        =   1429
      ButtonWidth     =   1085
      ButtonHeight    =   1376
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Load"
            Key             =   "Load"
            Object.ToolTipText     =   "Load"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Find"
            Key             =   "Find"
            Object.ToolTipText     =   "Find"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Restart"
            Key             =   "Restart"
            Object.ToolTipText     =   "Restart CDScan"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   5520
      TabIndex        =   12
      Top             =   4920
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ListBox List3 
      Height          =   255
      ItemData        =   "Form1.frx":5B22
      Left            =   6960
      List            =   "Form1.frx":5B24
      TabIndex        =   15
      Top             =   4920
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ListBox List4 
      Height          =   1230
      ItemData        =   "Form1.frx":5B26
      Left            =   240
      List            =   "Form1.frx":5B2D
      TabIndex        =   17
      Top             =   2520
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Disk Name"
      Height          =   195
      Left            =   7320
      TabIndex        =   13
      Top             =   960
      Width           =   780
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   3000
      TabIndex        =   9
      Top             =   5040
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuLoad 
         Caption         =   "&Load..."
      End
      Begin VB.Menu mnuFind 
         Caption         =   "&Find"
      End
      Begin VB.Menu mnuRestart 
         Caption         =   "&Restart"
      End
      Begin VB.Menu mnudiv1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About CDScan..."
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetFileTime Lib "kernel32" _
(ByVal hFile As Long, lpCreationTime As FILETIME, _
lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long


Public Sub LoadLV_Write()

Dim filename As String
Dim list As ListBox
Dim lstinput
Dim L2, L3, L4, L5, L6, L7, L8, L9 As String
Dim si As ListSubItem
Dim li As ListItem


filename = Form1.CommonDialog1.filename
    Open filename For Input As #1
    Do While Not EOF(1)
    Input #1, lstinput
    List6.AddItem lstinput
    
    Loop
    
    Close #1
    Dim i

    For i = 0 To List6.ListCount - 1
    L2 = Split(List6.list(i), Chr(149), -1)(0) 'disk
    L3 = Split(List6.list(i), Chr(149), -1)(1) 'path
    L4 = Split(List6.list(i), Chr(149), -1)(2) 'file
    L5 = Split(List6.list(i), Chr(149), -1)(3) 'size
    L6 = Split(List6.list(i), Chr(149), -1)(4) 'created
    L7 = Split(List6.list(i), Chr(149), -1)(5) 'modified
    L8 = Split(List6.list(i), Chr(149), -1)(6) 'attributes
    L9 = Split(List6.list(i), Chr(149), -1)(7) 'notes
    
    Set li = ListView1.ListItems.Add(1, , L2)
    Set si = li.ListSubItems.Add(1, , L3)
    Set si = li.ListSubItems.Add(2, , L4)
       si.ForeColor = vbRed
    Set si = li.ListSubItems.Add(3, , L5)
    Set si = li.ListSubItems.Add(4, , L6)
    Set si = li.ListSubItems.Add(5, , L7)
    Set si = li.ListSubItems.Add(6, , L8)
    Set si = li.ListSubItems.Add(7, , L9)
    

    Next i
   
StatusBar1.Panels(2) = ListView1.ListItems.Count & " Items"
End Sub
Public Sub LoadLV()
    Dim filename As String
    Dim list As ListBox
    Dim lstinput
    Dim L2, L3, L4, L5, L6, L7, L8, L9 As String
    Dim si As ListSubItem
    Dim li As ListItem
    
        ListView1.ListItems.Clear

filename = Form1.CommonDialog1.filename
    Open filename For Input As #1
    Do While Not EOF(1)
    Input #1, lstinput
    List6.AddItem lstinput
    
    Loop
    
    Close #1
    Dim i

    For i = 0 To List6.ListCount - 1
    L2 = Split(List6.list(i), Chr(149), -1)(0) 'disk
    L3 = Split(List6.list(i), Chr(149), -1)(1) 'path
    L4 = Split(List6.list(i), Chr(149), -1)(2) 'file
    L5 = Split(List6.list(i), Chr(149), -1)(3) 'size
    L6 = Split(List6.list(i), Chr(149), -1)(4) 'created
    L7 = Split(List6.list(i), Chr(149), -1)(5) 'modified
    L8 = Split(List6.list(i), Chr(149), -1)(6) 'attributes
    L9 = Split(List6.list(i), Chr(149), -1)(7) 'notes
    
    Set li = ListView1.ListItems.Add(1, , L2)
    Set si = li.ListSubItems.Add(1, , L3)
    Set si = li.ListSubItems.Add(2, , L4)
       si.ForeColor = vbRed
    Set si = li.ListSubItems.Add(3, , L5)
    Set si = li.ListSubItems.Add(4, , L6)
    Set si = li.ListSubItems.Add(5, , L7)
    Set si = li.ListSubItems.Add(6, , L8)
    Set si = li.ListSubItems.Add(7, , L9)
    

    Next i
   
StatusBar1.Panels(2) = ListView1.ListItems.Count & " Items"
End Sub


Public Sub WriteTmp(list As ListBox, filename As String)
    If List5.ListCount <= 0 Then
        MsgBox "WriteTmp Error - Listbox is empty - cannot write to file!", vbCritical, "Error"
    Exit Sub
    End If
    
    On Error GoTo ErrHandler
    
    'CommonDialog1.Flags = cdlOFNOverwritePrompt Or cdlOFNHideReadOnly Or cdlOFNPathMustExist
    'CommonDialog1.filename = Text3.Text
    'CommonDialog1.Filter = "CD Catalog Files (*.cdc)|*.cdc"
    'CommonDialog1.ShowSave
    'Open CommonDialog1.filename For Output As #1
    
    'filename1 = App.Path & "\" & "catdata.dat"
    Open filename For Output As #1
    
    For i = 0 To List7.ListCount - 1
    Dim fattr As String
    Dim tmp1, tmp2, tmp3, tmp4 As String
    tmp1 = "x"
    tmp2 = "x"
    tmp3 = "x"
    tmp4 = "x"
    
    'R-A-H-S
    If GetAttr(List7.list(e)) And FILE_ATTRIBUTE_READONLY Then tmp1 = "R"
    If GetAttr(List7.list(e)) And FILE_ATTRIBUTE_ARCHIVE Then tmp2 = "A"
    If GetAttr(List7.list(e)) And FILE_ATTRIBUTE_HIDDEN Then tmp3 = "H"
    If GetAttr(List7.list(e)) And FILE_ATTRIBUTE_SYSTEM Then tmp4 = "S"
   'If GetAttr(List4.list(e)) And FILE_ATTRIBUTE_NORMAL Then tmp = tmp & "NORMAL  "
   'If GetAttr(List4.list(e)) And FILE_ATTRIBUTE_TEMPORARY Then tmp = tmp & "TEMPORARY  "
   'If GetAttr(List4.list(e)) And FILE_ATTRIBUTE_COMPRESSED Then tmp = tmp & "COMPRESSED  "
   'If GetAttr(List4.list(e)) And FILE_ATTRIBUTE_DIRECTORY Then tmp = tmp & "DIRECTORY  "
    fattr = tmp1 & "-" & tmp2 & "-" & tmp3 & "-" & tmp4
    
        Write #1, Splitpath(List7.list(e), 4) & _
        Chr(149) & Splitpath(List7.list(e), 5) & _
        Chr(149) & Splitpath(List7.list(e), 6) & _
        Chr(149) & FileLen(List7.list(e)) & _
        Chr(149) & "created" & _
        Chr(149) & FileDateTime(List7.list(e)) & _
        Chr(149) & fattr & _
        Chr(149) & "add notes here... "
                
        List7.RemoveItem (e)
        
    Next
    
        Close #1
     Call Command5_Click
    'For i = 0 To List3.ListCount - 1
    '    Toolbar1.Buttons.Item(List3.list(i)).Enabled = False
    'Next i
    
    'Label2.Enabled = False
    'Text2.Enabled = False
    
    'command3.Enabled = False
    'command3.Visible = True
    
    'command4.Visible = True
    'command4.Enabled = False
    
    'command2.Visible = True
    'Exit Sub
ErrHandler:
   If err.Number = 32755 Then
   Exit Sub
   End If

If err.Number = 52 Then
err.Clear
Resume Next
ElseIf err.Number = 75 Then
err.Clear
Resume Next
ElseIf err.Number = 0 Then
err.Clear
'Resume Next
Else
    MsgBox "WriteTmp Error in WriteList" & Chr(13) & Chr(13) & err.Number _
    & " - " & err.Description, vbCritical, "Error"
    End If
    Exit Sub


End Sub

Public Function Splitpath(ByVal TempPath As String, ReturnType As Integer) As String



    Dim DriveLetter As String, DirPath As String, filename As String
    Dim Extension As String, PathLength As Integer
    Dim Offset As Integer, thisLength As Integer, FileNameFound As Boolean
    DriveLetter = ""
    DirPath = ""
    filename = ""
    Extension = ""


    If Mid(TempPath, 2, 1) = ":" Then ' Find the drive letter.
        DriveLetter = Left(TempPath, 2)
        TempPath = Mid(TempPath, 3)
    End If


    PathLength = Len(TempPath)


    For Offset = PathLength To 1 Step -1 ' Find the Next delimiter.


        Select Case Mid(TempPath, Offset, 1)
            Case ".": ' This indicates either an extension or a . or a ..
            thisLength = Len(TempPath$) - Offset%


            If thisLength >= 1 And thisLength <= 3 Then ' Extension
                Extension = Mid(TempPath, Offset, thisLength + 1)
            End If


            TempPath = Left(TempPath$, Offset - 1)
            Case "\": ' This indicates a path delimiter.
            thisLength = Len(TempPath$) - Offset


            If thisLength >= 1 And thisLength <= 40 Then ' Filename
                filename = Mid$(TempPath, Offset + 1, thisLength%)
                TempPath$ = Left(TempPath, Offset)
                DirPath = TempPath$
                FileNameFound = True
                Exit For
            End If


            Case Else
        End Select


Next Offset



If ReturnType = 1 Then
    Splitpath = TempPath & filename & Extension
ElseIf ReturnType = 2 Then
    Splitpath = DriveLetter & TempPath
ElseIf ReturnType = 3 Then
    Splitpath = Extension
ElseIf ReturnType = 4 Then
    Splitpath = DriveLetter
ElseIf ReturnType = 5 Then
    Splitpath = TempPath
ElseIf ReturnType = 6 Then
    Splitpath = filename & Extension
ElseIf ReturnType = 7 Then
    Splitpath = DriveLetter & TempPath & filename & Extension
End If


End Function


Private Sub AppendList(list As ListBox, filename As String)
If list.ListCount <= 0 Then
        MsgBox "Listbox is empty - cannot write to file!", vbCritical, "Error"
        Exit Sub
    End If
    
    On Error GoTo ErrHandler
    
    
    
    CommonDialog1.Filter = "All Files (*.*)|*.*|Text Files (*.txt)|*.txt|CD Catalog Files (*.cdc)|*.cdc"
    CommonDialog1.FilterIndex = 3
    CommonDialog1.ShowSave
   
   
    
    Open CommonDialog1.filename For Append As #1


    For i = 0 To list.ListCount - 1
    Dim fattr As String
    Dim tmp1, tmp2, tmp3, tmp4 As String
    tmp1 = "x"
    tmp2 = "x"
    tmp3 = "x"
    tmp4 = "x"
    
    'R-A-H-S
    If GetAttr(List5.list(e)) And FILE_ATTRIBUTE_READONLY Then tmp1 = "R"
    If GetAttr(List5.list(e)) And FILE_ATTRIBUTE_ARCHIVE Then tmp2 = "A"
    If GetAttr(List5.list(e)) And FILE_ATTRIBUTE_HIDDEN Then tmp3 = "H"
    If GetAttr(List5.list(e)) And FILE_ATTRIBUTE_SYSTEM Then tmp4 = "S"
   'If GetAttr(List4.list(e)) And FILE_ATTRIBUTE_NORMAL Then tmp = tmp & "NORMAL  "
   'If GetAttr(List4.list(e)) And FILE_ATTRIBUTE_TEMPORARY Then tmp = tmp & "TEMPORARY  "
   'If GetAttr(List4.list(e)) And FILE_ATTRIBUTE_COMPRESSED Then tmp = tmp & "COMPRESSED  "
   'If GetAttr(List4.list(e)) And FILE_ATTRIBUTE_DIRECTORY Then tmp = tmp & "DIRECTORY  "
    fattr = tmp1 & "-" & tmp2 & "-" & tmp3 & "-" & tmp4
    
   
    
        'For e = 0 To List4.ListCount - 1
        'Print #1, List.List(i) & "," & Text2.Text & ","
        'Write #1, List.List(i) & "," & Text2.Text
        'Write #1, List.List(i) & "," & Space(1) & Text2.Text,
        'Write #1, Text2.Text & Space(1) & "-" & Space(1) & List.List(I)
        
        'Write #1, Text2.Text & "," & Splitpath(List.List(I), 5) & "," & Splitpath(List.List(I), 6) _
        '& ",,,,," & "test"
        'Write #1, Text2.Text & Chr(149) & Splitpath(List.List(I), 5) & Chr(149) & Splitpath(List.List(I), 6) _
        '& Chr(149) & "size" & Chr(149) & "created" & Chr(149) & "modified" & Chr(149) & "attributes" & Chr(149) & "test"
        Write #1, Text2.Text & Chr(149) & Splitpath(list.list(i), 5) & _
        Chr(149) & Splitpath(list.list(i), 6) & _
        Chr(149) & FileLen(List5.list(e)) & _
        Chr(149) & "created" & _
        Chr(149) & FileDateTime(List5.list(e)) & _
        Chr(149) & fattr & _
        Chr(149) & "add notes here... "
                        
        'Debug.Print GetFileTime(List4.list(e), created1, AccessTime1, LastWriteTime1)
        
        List5.RemoveItem (e)
    Next
    
        Close #1
        
        List6.Clear
        ListView1.ListItems.Clear
        Call LoadLV
    
    For i = 0 To List3.ListCount - 1
        Toolbar1.Buttons.Item(List3.list(i)).Enabled = False
    Next i
    
    Label2.Enabled = False
    Text2.Enabled = False
    
    command4.Enabled = False
    command2.Enabled = True
    
    
    Exit Sub
ErrHandler:
    If err.Number = 32755 Then
   Exit Sub
   End If

If err.Number = 52 Then
err.Clear
Resume Next
ElseIf err.Number = 75 Then
err.Clear
Resume Next
Else
    MsgBox "Error in WriteList" & Chr(13) & Chr(13) & err.Number _
    & " - " & err.Description, vbCritical, "CDScan"
    End If
    Exit Sub
End Sub
Private Sub ScanCD(drive$)

    On Error GoTo ErrHandler
    
    List1.Clear
    List2.Clear
    List4.Clear
    List5.Clear
    List6.Clear
    
    ListView1.ListItems.Clear
    
    Dir1.Path = drive$
    File1.Path = Dir1.Path  'Set file path

'Add files in root directory to catalog

For t& = 0 To File1.ListCount - 1
    'List2.AddItem Dir1.Path + File1.List(t&)
    List2.AddItem "\" + File1.list(t&)
    'List4.AddItem FileLen(File1.Path + File1.list(t&))
    List5.AddItem File1.Path + File1.list(t&)
    List7.AddItem File1.Path + File1.list(t&)
    Next

'Add initial subdirectories in root
    For t& = 0 To Dir1.ListCount - 1
        List1.AddItem Dir1.list(t&)
    Next

    If t& = 0 Then GoTo finishedpoint 'No subdirectories, we're done.

    cnt& = 0  'Start with first directory in list

startpoint:

    Dir1.Path = List1.list(cnt&)
    File1.Path = Dir1.Path

'Add any subdirectories found

    For t& = 0 To Dir1.ListCount - 1
        List1.AddItem Dir1.list(t&)
    Next

'Now add files found...
    For t& = 0 To File1.ListCount - 1
        'List2.AddItem Dir1.Path + "\" + File1.List(t&)
        'List2.AddItem "\" + File1.List(t&)
        'List2.AddItem Dir1.Path + File1.List(t&)
        List2.AddItem Splitpath(Dir1.Path, 1) + "\" + File1.list(t&)
        'List4.AddItem FileLen(Splitpath(Dir1.Path, 7) + "\" + File1.list(t&))
        List5.AddItem Splitpath(Dir1.Path, 7) + "\" + File1.list(t&)
        List7.AddItem Splitpath(Dir1.Path, 7) + "\" + File1.list(t&)
    Next

'Increment count
        cnt& = cnt& + 1
        If cnt& < List1.ListCount Then GoTo startpoint


finishedpoint:
    If List2.ListCount > 0 Then
        Label2.Enabled = True
        Text2.Enabled = True
        command3.Enabled = True
    End If
'Label1.Caption = "Total:" + Str$(List2.ListCount) & Space(1) & "files"
    StatusBar1.Panels.Item(1).Text = "Total:" + Str$(List2.ListCount) & Space(1) & "files"
'MsgBox List4.ListCount

ErrHandler:
     If err.Number = 68 Then
          MsgBox "This Drive is unavailable Please insert a CD or Close the CD Drive Door", vbCritical, "CDScan"
     Exit Sub
     
     End If
     Call WriteTmp(List2, App.Path & "\" & "catdata.dat")
End Sub
Private Sub cmdCountAllFiles_Click()

   Dim sSourcePath As String
   Dim sDestination As String
   Dim sFileType As String
   Dim numFiles As Long
      
  'set the appropriate initializing values
   sSourcePath = "c:\win\system\"
   sFileType = "*.*"
      
  'get the count
   numFiles = rgbCountFilesAll(sSourcePath, sFileType)
   MsgBox numFiles & " files found matching " & sSourcePath & sFileType

End Sub


Private Sub cmdCountByAttribute_Click()

   Dim sSourcePath As String
   Dim sDestination As String
   Dim sFileType As String
   Dim sFileAttributes As Long
   Dim numFiles As Long
      
  'set the appropriate initializing values
   sSourcePath = "c:\win\"
   sFileType = "*.*"
      
  'set the attribute(s) to search for
   sFileAttributes = FILE_ATTRIBUTE_NORMAL Or FILE_ATTRIBUTE_ARCHIVE
      
  'get the count
   numFiles = rgbCountFilesByAttribute(sSourcePath, sFileType, sFileAttributes)
   MsgBox numFiles & " files found matching " & _
          sSourcePath & sFileType & " with attribute(s) " & _
          GetAttributeString(sFileAttributes)

End Sub


Private Sub cmdListFilesByAttribute_Click()

   Dim sSourcePath As String
   Dim sDestination As String
   Dim sFileType As String
   Dim sFileAttributes As Long
   Dim numFiles As Long
   Dim r As Long
      
  'set the appropriate initializing values
   sSourcePath = "c:\win\"
   sFileType = "*.*"
      
  'set the attribute(s) to search for
   sFileAttributes = FILE_ATTRIBUTE_NORMAL Or FILE_ATTRIBUTE_ARCHIVE
      
  'clear the list and lock its window to
  'prevent updating until the search is complete
   List1.Clear
   r = LockWindowUpdate(List1.hWnd)
   DoEvents
      
  'get the list and count
   numFiles = rgbListFilesByAttribute(sSourcePath, sFileType, sFileAttributes)
      
  'free the locked window
   r = LockWindowUpdate(0)
   
   MsgBox numFiles & " files found matching " & _
          sSourcePath & sFileType & " with attribute(s) " & _
          GetAttributeString(sFileAttributes), vbOKOnly, "File Count"

End Sub


Public Function rgbCountFilesAll(sSourcePath As String, _
                                 sFileType As String) As Long

   Dim WFD As WIN32_FIND_DATA
   Dim SA As SECURITY_ATTRIBUTES
   
   Dim r As Long
   Dim hFile As Long
   Dim bNext As Long
   Dim fCount As Long
   Dim currFile As String
      
  'Start searching for files in the Source directory by
  'obtaining a file handle to the first file matching the
  'filespec passed
   hFile = FindFirstFile(sSourcePath & sFileType, WFD)
   
   If (hFile = INVALID_HANDLE_VALUE) Then
     
      'no match, so bail out
       rgbCountFilesAll = 0
       Exit Function
       
   End If
       
  'must have at least one, so ...
   If hFile Then
      
      Do
         
        'increment the counter and find the next
        'file matching the filespec
         fCount = fCount + 1
         bNext = FindNextFile(hFile, WFD)
               
      Loop Until bNext = 0
      
   End If
      
  'Close the search handle
   r = FindClose(hFile)
      
  'return the number of files found
   rgbCountFilesAll = fCount
   
End Function


Public Function rgbCountFilesByAttribute(sSourcePath As String, _
                                         sFileType As String, _
                                         AttrFlags As Long) As Long

   Dim WFD As WIN32_FIND_DATA
   Dim SA As SECURITY_ATTRIBUTES
   
   Dim r As Long
   Dim hFile As Long
   Dim bNext As Long
   Dim fCount As Long
   Dim currFile As String
      
  'Start searching for files in the Source directory.
   hFile = FindFirstFile(sSourcePath & sFileType, WFD)
   
   If (hFile = INVALID_HANDLE_VALUE) Then
   
       rgbCountFilesByAttribute = 0
       Exit Function
       
   End If
    
   If hFile Then
      
      Do
         
        'trim trailing nulls, leaving one to terminate the string
         currFile = Left$(WFD.cFileName, InStr(WFD.cFileName, Chr$(0)))
            
        'compare the file attributes against the passes flag(s)
         If AttrFlags And GetFileAttributes(sSourcePath & currFile) Then
            
           'increment the counter and find the next
           'file matching the initial file spec
            fCount = fCount + 1
            
         End If
         
         bNext = FindNextFile(hFile, WFD)
               
      Loop Until bNext = 0
      
   End If
      
  'Close the search handle
   r = FindClose(hFile)
      
  'return the number of files
   rgbCountFilesByAttribute = fCount
   
End Function


Public Function rgbListFilesByAttribute(sSourcePath As String, _
                                        sFileType As String, _
                                        AttrFlags As Long) As Long

   Dim WFD As WIN32_FIND_DATA
   Dim SA As SECURITY_ATTRIBUTES
   
   Dim r As Long
   Dim hFile As Long
   Dim bNext As Long
   Dim fCount As Long
   Dim currFile As String
      
  'Start searching for files in the Source directory.
   hFile = FindFirstFile(sSourcePath & sFileType, WFD)
   
   If (hFile = INVALID_HANDLE_VALUE) Then
   
       rgbListFilesByAttribute = 0
       Exit Function
       
   End If
    
   If hFile Then
      
      Do
         
        'trim trailing nulls, leaving one to terminate the string
         currFile = Left$(WFD.cFileName, InStr(WFD.cFileName, Chr$(0)) - 1)
            
        'compare the file attributes against the passes flag(s)
         If AttrFlags And GetFileAttributes(sSourcePath & currFile) Then
             
           'increment the counter and find the next
           'file matching the initial file spec
            List1.AddItem currFile
            fCount = fCount + 1
            
         End If
         
         bNext = FindNextFile(hFile, WFD)
               
      Loop Until bNext = 0
      
   End If
      
  'Close the search handle
   r = FindClose(hFile)
      
  'return the number of files fCount
   rgbListFilesByAttribute = fCount
   
End Function


Public Function GetAttributeString(attr As Long) As String

   Dim tmp As String
   
   If attr And FILE_ATTRIBUTE_ARCHIVE Then tmp = tmp & "ARCHIVE  "
   If attr And FILE_ATTRIBUTE_NORMAL Then tmp = tmp & "NORMAL  "
   If attr And FILE_ATTRIBUTE_HIDDEN Then tmp = tmp & "HIDDEN  "
   If attr And FILE_ATTRIBUTE_READONLY Then tmp = tmp & "READONLY  "
   If attr And FILE_ATTRIBUTE_SYSTEM Then tmp = tmp & "SYSTEM  "
   If attr And FILE_ATTRIBUTE_TEMPORARY Then tmp = tmp & "TEMPORARY  "
   If attr And FILE_ATTRIBUTE_COMPRESSED Then tmp = tmp & "COMPRESSED  "
   If attr And FILE_ATTRIBUTE_DIRECTORY Then tmp = tmp & "DIRECTORY  "

   GetAttributeString = tmp

End Function
Public Sub ReadList(list As ListBox, filename As String, Optional ClearList As Boolean)



    On Error GoTo err
    Open filename For Input As #1
    If ClearList = True Then list.Clear

    Do While Not EOF(1)
        Input #1, lstinput
        list.AddItem lstinput
        'List.AddItem Right$(lstinput, 11)  'need to parse string so it only loads the file name and disk#
    Loop

    Close #1
    Call LoadLV
    Exit Sub
err:
        
    MsgBox "Error Reading file" & Chr(13) & Chr(13) & err.Number _
    & " - " & err.Description, vbCritical, "CDScan"
    
    Exit Sub
End Sub



Public Sub WriteList(list As ListBox, filename As String)

        
      
        
    If list.ListCount <= 0 Then
        MsgBox "Listbox is empty - cannot write to file!", vbCritical, "Error"
        Exit Sub
    End If
    
    On Error GoTo ErrHandler
    
    CommonDialog1.Flags = cdlOFNOverwritePrompt Or cdlOFNHideReadOnly Or cdlOFNPathMustExist
    CommonDialog1.filename = Text3.Text
    CommonDialog1.Filter = "CD Catalog Files (*.cdc)|*.cdc"
    CommonDialog1.ShowSave
    Open CommonDialog1.filename For Output As #1
    
    
    For i = 0 To list.ListCount - 1
    Dim fattr As String
    Dim tmp1, tmp2, tmp3, tmp4 As String
    tmp1 = "x"
    tmp2 = "x"
    tmp3 = "x"
    tmp4 = "x"
    
    'R-A-H-S
    If GetAttr(List5.list(e)) And FILE_ATTRIBUTE_READONLY Then tmp1 = "R"
    If GetAttr(List5.list(e)) And FILE_ATTRIBUTE_ARCHIVE Then tmp2 = "A"
    If GetAttr(List5.list(e)) And FILE_ATTRIBUTE_HIDDEN Then tmp3 = "H"
    If GetAttr(List5.list(e)) And FILE_ATTRIBUTE_SYSTEM Then tmp4 = "S"
   'If GetAttr(List4.list(e)) And FILE_ATTRIBUTE_NORMAL Then tmp = tmp & "NORMAL  "
   'If GetAttr(List4.list(e)) And FILE_ATTRIBUTE_TEMPORARY Then tmp = tmp & "TEMPORARY  "
   'If GetAttr(List4.list(e)) And FILE_ATTRIBUTE_COMPRESSED Then tmp = tmp & "COMPRESSED  "
   'If GetAttr(List4.list(e)) And FILE_ATTRIBUTE_DIRECTORY Then tmp = tmp & "DIRECTORY  "
    fattr = tmp1 & "-" & tmp2 & "-" & tmp3 & "-" & tmp4
    
   
    
        'For e = 0 To List4.ListCount - 1
        'Print #1, List.List(i) & "," & Text2.Text & ","
        'Write #1, List.List(i) & "," & Text2.Text
        'Write #1, List.List(i) & "," & Space(1) & Text2.Text,
        'Write #1, Text2.Text & Space(1) & "-" & Space(1) & List.List(I)
        
        'Write #1, Text2.Text & "," & Splitpath(List.List(I), 5) & "," & Splitpath(List.List(I), 6) _
        '& ",,,,," & "test"
        'Write #1, Text2.Text & Chr(149) & Splitpath(List.List(I), 5) & Chr(149) & Splitpath(List.List(I), 6) _
        '& Chr(149) & "size" & Chr(149) & "created" & Chr(149) & "modified" & Chr(149) & "attributes" & Chr(149) & "test"
        Write #1, Text2.Text & Chr(149) & Splitpath(list.list(i), 5) & _
        Chr(149) & Splitpath(list.list(i), 6) & _
        Chr(149) & FileLen(List5.list(e)) & _
        Chr(149) & "created" & _
        Chr(149) & FileDateTime(List5.list(e)) & _
        Chr(149) & fattr & _
        Chr(149) & "add notes here... "
                        
        'Debug.Print GetFileTime(List4.list(e), created1, AccessTime1, LastWriteTime1)
        
        List5.RemoveItem (e)
    Next
    
        Close #1
        
      
        'For lci = 0 To List6.ListCount - 1
         'MsgBox List6.ListCount
        '    List6.RemoveItem lci
        'Next lci
        List6.Clear
        ListView1.ListItems.Clear
        Call LoadLV
        'Call ReadList(Form1.List2, Form1.CommonDialog1.filename, True)
        
    For i = 0 To List3.ListCount - 1
        Toolbar1.Buttons.Item(List3.list(i)).Enabled = False
    Next i
    
    Label2.Enabled = False
    Text2.Enabled = False
    
    command3.Enabled = False
    command3.Visible = True
    
    command4.Visible = True
    command4.Enabled = False
    
    command2.Visible = True
    Exit Sub
ErrHandler:
   If err.Number = 32755 Then
   Exit Sub
   End If

If err.Number = 52 Then
err.Clear
Resume Next
ElseIf err.Number = 75 Then
err.Clear
Resume Next
Else
    MsgBox "Error in WriteList" & Chr(13) & Chr(13) & err.Number _
    & " - " & err.Description, vbCritical, "Error"
    End If
    Exit Sub
End Sub



Private Sub Command1_Click()
'Call ReadList(List2, "C:\test.list", True)  'Will input contents of c:\test.list to the listbox (List1) and WILL clear the list beforehand
End Sub


Private Sub Command2_Click()
MsgBox "Place The next CD in the CD-Drive and click OK", vbInformation, "CDScan"
List2.Clear
Text2.Text = ""
StatusBar1.Panels.Item(1).Text = ""
'For i = 0 To List3.ListCount ''dont touch
'Toolbar1.Buttons.Item(List3.List(0)).Enabled = True 'dont touch
For i = 0 To List3.ListCount - 1
Toolbar1.Buttons.Item(List3.list(i)).Enabled = True
Next i

command3.Visible = False
command2.Enabled = False
command4.Enabled = True
'If List2.ListCount <= 0 Then
'Command2.Visible = False
'End If
End Sub

Private Sub Command3_Click()
'Call WriteList(List2, "C:\list.txt")
If Text2.Text = "" Then
MsgBox "Please enter a Disk Name or Number", vbExclamation, "CDScan"
Exit Sub
End If
Call WriteList(List2, CommonDialog1.filename)
End Sub

Private Sub Command4_Click()
If Text2.Text = "" Then
MsgBox "Please enter a Disk Name or Number", vbExclamation, "CDScan"
Exit Sub
End If
Call AppendList(List2, CommonDialog1.filename)
End Sub



Private Sub Command5_Click()
Dim filename As String
Dim list As ListBox
Dim lstinput
Dim L2, L3, L4, L5, L6, L7, L8, L9 As String
Dim si As ListSubItem
Dim li As ListItem

'List7.Clear

filename = App.Path & "\" & "catdata.dat"
    Open filename For Input As #1
    Do While Not EOF(1)
    Input #1, lstinput
    List6.AddItem lstinput
    
    Loop
    
    Close #1
    Dim i

    For i = 0 To List6.ListCount - 1
    L2 = Split(List6.list(i), Chr(149), -1)(0) 'disk
    L3 = Split(List6.list(i), Chr(149), -1)(1) 'path
    L4 = Split(List6.list(i), Chr(149), -1)(2) 'file
    L5 = Split(List6.list(i), Chr(149), -1)(3) 'size
    L6 = Split(List6.list(i), Chr(149), -1)(4) 'created
    L7 = Split(List6.list(i), Chr(149), -1)(5) 'modified
    L8 = Split(List6.list(i), Chr(149), -1)(6) 'attributes
    L9 = Split(List6.list(i), Chr(149), -1)(7) 'notes
    'L5 = Split(List1.list(i), ",", -1)(3) 'size
    Set li = ListView1.ListItems.Add(1, , L2)
    Set si = li.ListSubItems.Add(1, , L3)
    Set si = li.ListSubItems.Add(2, , L4)
       si.ForeColor = vbRed
    Set si = li.ListSubItems.Add(3, , L5)
    Set si = li.ListSubItems.Add(4, , L6)
    Set si = li.ListSubItems.Add(5, , L7)
    Set si = li.ListSubItems.Add(6, , L8)
    Set si = li.ListSubItems.Add(7, , L9)
    

    Next i
   
   Kill App.Path & "\" & "catdata.dat"
StatusBar1.Panels(2) = ListView1.ListItems.Count & " Items"
End Sub



Private Sub Form_Load()

   Dim sngWidth As Single
   Dim si As ListSubItem
   Dim li As ListItem

   ' You can't see ColumnHeaders or ListSubitems
   ' unless the View is set to lvwReport.
   ListView1.View = lvwReport
   ListView1.AllowColumnReorder = True
   ' Calculate the width of a ColumnHeader object.
   sngWidth = ListView1.Width / 5
   
   ListView1.ColumnHeaders.Add 1, Text:="Disk", Width:=1000
   ListView1.ColumnHeaders.Add 2, Text:="Path", Width:=sngWidth
   ListView1.ColumnHeaders.Add 3, Text:="Filename", Width:=sngWidth
   ListView1.ColumnHeaders.Add 4, Text:="Size", Width:=sngWidth
   ListView1.ColumnHeaders.Add 5, Text:="Created", Width:=sngWidth
   ListView1.ColumnHeaders.Add 6, Text:="Modified", Width:=sngWidth
   ListView1.ColumnHeaders.Add 7, Text:="Attributes", Width:=sngWidth
   ListView1.ColumnHeaders.Add 8, Text:="Notes", Width:=4000
   
        
    If List2.ListCount > 0 Then
        Label2.Enabled = True
        command3.Enabled = True
    End If
        
    
    command3.Visible = True
    command2.Visible = False
    
 alldrives$ = Space$(256)
       

ret& = GetLogicalDriveStrings(Len(alldrives$), alldrives$)

'trim off any trailing spaces. AllDrives$
'now contains all the drive letters.
alldrives$ = Left$(alldrives$, ret&)

Do
   'first check that there is a chr$(0) in the string
   pos% = InStr(alldrives$, Chr$(0))
   'if there's one, then...

     If pos% Then
     'extract the drive up to the chr$(0)
     justonedrive$ = Left$(alldrives$, pos% - 1)
               
     'and remove that from the Alldrives string,
     'so it won't be checked again
     alldrives$ = Mid$(alldrives$, pos% + 1, Len(alldrives$))
               
     'with the one drive, call the API to
     'determine the drive type
     DriveType& = GetDriveType(justonedrive$)
     'check if it's what we want

   Dim btnbutton As Button
   
   
             If DriveType& = 5 Then 'then it is a CD Drive
             
            Set btnbutton = Toolbar1.Buttons.Add(, (justonedrive$), (UCase$(justonedrive$)), , 4)
            'Toolbar1. = justonedrive$
            List3.AddItem (justonedrive$)
            
            For i = 0 To List3.ListCount - 1
                Toolbar1.Buttons.Item(List3.list(i)).ToolTipText = "Scan CD"
            Next i
            
            
            
            If Form2.Option1.Value = True Then
            Toolbar1.Buttons.Item(1).Enabled = False
            'Toolbar1.Buttons.Item(2).Enabled = False
            mnuLoad.Enabled = False
            'Toolbar1.Buttons.Item(3).Enabled = False
            
            End If
            
            If Form2.Option3.Value = True Then
            'Toolbar1.Buttons.Item(1).Enabled = False
            Toolbar1.Buttons.Item(3).Enabled = True
            Toolbar1.Buttons.Item(justonedrive$).Enabled = False
            command3.Enabled = False
            command4.Enabled = False
            command2.Enabled = False
            
            End If
            
            If Form2.Option4.Value = True Then
            command3.Enabled = False
            command4.Enabled = False
            command2.Visible = True
            'For i = 0 To justonedrive$
            Toolbar1.Buttons.Item(justonedrive$).Enabled = False
            'Next i
            End If
            
                '''Print UCase$(JustOneDrive$) & " is a CD Drive"
             'Else
             '   Print UCase$(JustOneDrive$) & " is NOT a CD Drive"
             End If

     End If

Loop Until alldrives$ = ""

StatusBar1.Panels(2) = ListView1.ListItems.Count & " Items"

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

   'If Toolbar1.Buttons.Item(3).
   '      Unload Form1
   '      Form2.Show
   ' End If
        
   ' If Form3.Visible = True Then
   '      Unload Form
   '     Exit Sub
   '     End If
   ' End If
    
    'End
    
End Sub

Private Sub Form_Resize()
On Error Resume Next
If Form1.WindowState = 1 Then
Exit Sub
End If
If Form1.Height < 1700 Then
Exit Sub
End If
ListView1.Width = Form1.Width - 1400
'fix so it doesnt error on resizing below 1150
ListView1.Height = Form1.Height - 1750

command3.Left = ListView1.Width + 50
command2.Left = ListView1.Width + 50
command4.Left = ListView1.Width + 50
Label2.Left = ListView1.Width + 50
Text2.Left = ListView1.Width + 50
'MsgBox command3.Left
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
ListView1.SortKey = ColumnHeader.Index - 1
'ListView1.SortOrder = lvwDescending
    ListView1.Sorted = True
End Sub

Private Sub mnuAbout_Click()
Form4.Show vbModal
End Sub



Private Sub mnuExit_Click()
Unload Me
End
End Sub



Private Sub mnuFind_Click()
Form3.Show vbModal
End Sub

Private Sub mnuLoad_Click()
CommonDialog1.Flags = cdlOFNOverwritePrompt Or cdlOFNHideReadOnly Or cdlOFNPathMustExist
        CommonDialog1.Filter = "CD Catalog Files (*.cdc)|*.cdc"
        CommonDialog1.ShowOpen
        Text3.Text = Form1.CommonDialog1.filename
        Call ReadList(List2, CommonDialog1.filename, True)
        
    Form1.Caption = CommonDialog1.filename & Space(1) & "-" & Space(1) & "CDScan"
End Sub

Private Sub mnuRestart_Click()
Form2.Show
      Unload Form3
     Unload Form1
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo ErrHandler

   Dim sSourcePath As String
   Dim sDestination As String
   Dim sFileType As String
   Dim sFileAttributes As Long
   Dim numFiles As Long
   Dim r As Long
   
Select Case Button.Key
   Case Is = "Load"       ' Open file.
     ' MsgBox "Loading!"
        CommonDialog1.Flags = cdlOFNOverwritePrompt Or cdlOFNHideReadOnly Or cdlOFNPathMustExist
        CommonDialog1.Filter = "CD Catalog Files (*.cdc)|*.cdc"
        CommonDialog1.ShowOpen
        Text3.Text = Form1.CommonDialog1.filename
        Call ReadList(List2, CommonDialog1.filename, True)
        Form1.Caption = CommonDialog1.filename & Space(1) & "-" & Space(1) & "CDScan"
        
   Case Is = "Find"        ' Save file.
        Form3.Show vbModal
   Case Is = "Restart"
      'MsgBox "Deleting"
      Form2.Show
      Unload Form3
     Unload Form1
      
   Case Is = "a:\"
        Call ScanCD("a:\")
   Case Is = "b:\"
        Call ScanCD("b:\")
   Case Is = "c:\"
        Call ScanCD("c:\")
   Case Is = "d:\"
        Call ScanCD("d:\")
   Case Is = "e:\"
        Call ScanCD("e:\")
   Case Is = "f:\"
        Call ScanCD("f:\")
   Case Is = "g:\"
        Call ScanCD("g:\")
   Case Is = "h:\"
        Call ScanCD("h:\")
   Case Is = "i:\"
        Call ScanCD("i:\")
   Case Is = "j:\"
        Call ScanCD("j:\")
   Case Is = "k:\"
        Call ScanCD("k:\")
   Case Is = "l:\"
        Call ScanCD("l:\")
   Case Is = "m:\"
        Call ScanCD("m:\")
   Case Is = "n:\"
        Call ScanCD("n:\")
   Case Is = "o:\"
        Call ScanCD("o:\")
   Case Is = "p:\"
        Call ScanCD("p:\")
   Case Is = "q:\"
        Call ScanCD("q:\")
   Case Is = "r:\"
        Call ScanCD("r:\")
   Case Is = "s:\"
        Call ScanCD("s:\")
   Case Is = "t:\"
        Call ScanCD("t:\")
   Case Is = "u:\"
        Call ScanCD("u:\")
   Case Is = "v:\"
        Call ScanCD("v:\")
   Case Is = "w:\"
        Call ScanCD("w:\")
   Case Is = "x:\"
        Call ScanCD("x:\")
    Case Is = "y:\"
        Call ScanCD("y:\")
   Case Is = "z:\"
        Call ScanCD("z:\")
   End Select
ErrHandler:
    If err.Number = 32755 Then
   Exit Sub
   End If

End Sub

