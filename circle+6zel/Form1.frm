VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Form1"
   ClientHeight    =   5175
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7800
   DrawWidth       =   3
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   345
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   520
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "run"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
ScaleMode = 3
X = 200
y = 200
r = 70
t = 50

For a = 0 To 2 * 3.14 Step 2 * (3.14 / 6)
    xc = X + r * Cos(a)
    yc = y - r * Sin(a)
    PSet (xc, yc), vbBlack
    Circle (xc, yc), 20
Next a
Circle (X, y), 100, vbBlack
CurrentX = X + t * Cos(3.14 / 6)
CurrentY = y + t * Sin(3.14 / 6)
For a = (3.14 / 6) To 2 * 3.14 Step 2 * (3.14 / 6)
    
    xc = X + t * Cos(a)
    yc = y - t * Sin(a)
    
    Line -(xc, yc), vbBlack
Next a

End Sub
