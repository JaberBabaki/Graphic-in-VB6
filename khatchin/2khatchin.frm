VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5130
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7905
   LinkTopic       =   "Form1"
   ScaleHeight     =   342
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   527
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "cls"
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1200
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1200
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   1200
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "chin"
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label label3 
      Alignment       =   2  'Center
      Caption         =   "y"
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label label2 
      Alignment       =   2  'Center
      Caption         =   "x2"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   975
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Caption         =   "x1"
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text1.SetFocus
Form1.Cls
End Sub

Private Sub Command3_Click()
x1 = Text1.Text
X2 = Text2.Text
Y = Text3.Text
If x1 > X2 Then
 t = x1
 x1 = X2
 X2 = t
 End If
For i = x1 To X2
 If i Mod 10 = 0 Then
 i = i + 5
 End If
 PSet (i, Y), vbBlue
Next i

End Sub

