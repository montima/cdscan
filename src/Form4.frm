VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6300
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   6300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Left            =   4080
      TabIndex        =   2
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "mmonti@mminternet.com"
      Height          =   255
      Left            =   3240
      TabIndex        =   6
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "For comments or suggestions on CDScan email:"
      Height          =   495
      Left            =   3000
      TabIndex        =   5
      Top             =   1920
      Width           =   2775
   End
   Begin VB.Label Label4 
      Caption         =   "http://home.mminternet.com/~mmonti/CDScan"
      Height          =   255
      Left            =   2640
      TabIndex        =   4
      Top             =   1440
      Width           =   3375
   End
   Begin VB.Label Label3 
      Caption         =   "Copyright © 1999-2000"
      Height          =   255
      Left            =   2640
      TabIndex        =   3
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Created by Michael Monti"
      Height          =   255
      Left            =   2640
      TabIndex        =   1
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "CDScan v1.0b"
      Height          =   255
      Left            =   2640
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   3750
      Left            =   120
      Picture         =   "Form4.frx":1CFA
      Top             =   120
      Width           =   2250
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Label4.ForeColor = vbBlue Then
        Label4.ForeColor = vbBlack
    End If
        If Label6.ForeColor = vbBlue Then
         Label6.ForeColor = vbBlack
    End If
End Sub

Private Sub Label4_Click()
 'Shell ("http://www.planet-cource-code.com"), vbNormalNoFocus
 
 Dim ret
 ret = Shell("start http://home.mminternet.com/~mmonti/cdscan.htm", vbHide)
Unload Form4
 
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = vbBlue
End Sub

Private Sub Label6_Click()
Dim ret
 ret = Shell("start mailto:mmonti@mminternet.com", vbHide)
 Unload Form4
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.ForeColor = vbBlue
End Sub
