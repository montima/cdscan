VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CDScan"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6855
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "What do you want to do?"
      Height          =   3735
      Left            =   2520
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin VB.OptionButton Option4 
         Caption         =   "Add a new Disk to an Existing CD Catalog"
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   1560
         Width           =   3615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   480
         TabIndex        =   6
         Top             =   2280
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cancel"
         Height          =   495
         Left            =   2640
         TabIndex        =   5
         Top             =   3000
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   495
         Left            =   360
         TabIndex        =   4
         Top             =   3000
         Width           =   1215
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Search an Existing CD Catalog"
         Height          =   255
         Left            =   480
         TabIndex        =   3
         Top             =   1200
         Width           =   3615
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Open your last CD Catalog"
         Enabled         =   0   'False
         Height          =   255
         Left            =   480
         TabIndex        =   2
         Top             =   840
         Width           =   3615
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Create a new CD Catalog"
         Height          =   255
         Left            =   480
         TabIndex        =   1
         Top             =   480
         Value           =   -1  'True
         Width           =   3615
      End
      Begin VB.Label Label1 
         Caption         =   "Type a name for your new Catalog"
         Height          =   255
         Left            =   480
         TabIndex        =   7
         Top             =   2040
         Visible         =   0   'False
         Width           =   3135
      End
   End
   Begin VB.Image Image1 
      Height          =   3750
      Left            =   120
      Picture         =   "Form2.frx":1CFA
      Top             =   120
      Width           =   2250
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

 On Error GoTo ErrHandler
    If Option1.Value = True Then '//// option1 open new
       If Text1.Text = "" Then
         MsgBox " Please enter a name for the new catalog", vbExclamation, "CDScan"
         Exit Sub
        End If
         Form1.Text3.Text = Text1.Text
         Form1.Caption = Text1.Text & ".cdc" & Space(1) & "-" & Space(1) & "CDScan"
         Unload Form2
        Form1.Label2.Enabled = False
        Form1.Text2.Enabled = False
        Form1.Command3.Enabled = False
        Form1.Command4.Visible = False
        Form1.Show
    ElseIf Option2.Value = True Then '//// option2 open the last open catalog
        MsgBox "option2"
        Form1.Show
        Unload Form2
    ElseIf Option3.Value = True Then
         '//// option3 open an existing
        Form1.CommonDialog1.Flags = cdlOFNOverwritePrompt Or cdlOFNHideReadOnly Or cdlOFNPathMustExist
        Form1.CommonDialog1.Filter = "CD Catalog Files (*.cdc)|*.cdc"
        Form1.CommonDialog1.ShowOpen
        Form1.Text3.Text = Form1.CommonDialog1.Filename
        
    
        
        Call Form1.ReadList(Form1.List2, Form1.CommonDialog1.Filename, True)
        
        Form1.Caption = Form1.CommonDialog1.Filename & Space(1) & "-" & Space(1) & "CDScan"
        
        Form1.Show
        
        Unload Form2
        
    Else
        Form1.CommonDialog1.Flags = cdlOFNOverwritePrompt Or cdlOFNHideReadOnly Or cdlOFNPathMustExist
        Form1.CommonDialog1.Filter = "CD Catalog Files (*.cdc)|*.cdc"
        Form1.CommonDialog1.ShowOpen
        Form1.Text3.Text = Form1.CommonDialog1.Filename
        
        Call Form1.ReadList(Form1.List2, Form1.CommonDialog1.Filename, True)
        
        Form1.Caption = Form1.CommonDialog1.Filename & Space(1) & "-" & Space(1) & "CDScan"
        
    Form1.Command3.Enabled = False
    Form1.Command3.Visible = True
    
    Form1.Command4.Visible = True
    Form1.Command4.Enabled = False
    
    Form1.Command2.Visible = True
    Form1.Text2.Enabled = False
    Form1.Label2.Enabled = False
        Form1.Show
        
        Unload Form2
        
    End If
    
ErrHandler:
 Exit Sub
    
End Sub

Private Sub Command2_Click()
    'Form1.Show
    Unload Form2
    End
End Sub

Private Sub Form_Load()
        
        Label1.Visible = True
        Text1.Visible = True
        'Text1.CausesValidation = False
        Text1.Text = "Enter Catalog name here"
       
        
        
        
        
   
End Sub

Private Sub Option1_Click()
        Label1.Visible = True
        Text1.Visible = True
End Sub

Private Sub Option2_Click()
        Label1.Visible = False
        Text1.Visible = False
End Sub

Private Sub Option3_Click()
        Label1.Visible = False
        Text1.Visible = False
End Sub

Private Sub Option4_Click()
        Label1.Visible = False
        Text1.Visible = False
End Sub
