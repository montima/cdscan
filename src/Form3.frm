VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4545
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   4545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Results >>"
      Height          =   495
      Left            =   3120
      TabIndex        =   4
      Top             =   720
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   240
      TabIndex        =   5
      Top             =   1800
      Width           =   4095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   1680
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Find "
      Default         =   -1  'True
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Close <<"
      Height          =   495
      Left            =   3120
      TabIndex        =   6
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Search Results"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Find what?"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   780
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

List1.Clear

If Text1.Text = "" Then
MsgBox "Enter a search string", vbCritical, "CDScan"
Exit Sub
End If

              For I = 0 To Form1.List2.ListCount - 1

                If InStr(1, Form1.List2.List(I), Trim(Text1.Text), vbTextCompare) Then
            
            
            List1.AddItem (Form1.List2.List(I))
                
                
            
            End If

                
       Next I

        If List1.ListCount > 0 Then
         MsgBox "Found" & Space(1) & List1.ListCount & Space(1) & "Item(s)", vbInformation, "CDScan"
         End If
        
        If List1.ListCount = 0 Then
        MsgBox "Search text not found", vbExclamation, "CDScan"
        
    End If

End Sub

Private Sub Command2_Click()


Unload Me
End Sub

Private Sub Command3_Click()
Form3.Height = 4110
Command3.Visible = False
Command4.Visible = True


End Sub

Private Sub Command4_Click()
    Form3.Height = 1785
    Command4.Visible = False
    Command3.Visible = True
End Sub

Private Sub Form_Load()
Form3.Height = 1785
Text1.Text = ""

End Sub

