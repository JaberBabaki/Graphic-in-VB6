VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "run"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim t As Boolean
Private Sub Command1_Click()
Cls
ScaleMode = 3
xc1 = 300: Yc1 = 300: R1 = 100
Call Polygan(xc1, Yc1, R1, 16)

End Sub

Private Sub Polygan(ByVal Xc As Integer, ByVal Yc As Integer, ByVal R As Integer, ByVal N As Integer)
CurrentX = Xc + R * Cos(3.14 / 6)
CurrentY = Yc - R * Sin(3.14 / 6)
For Alpha = 3.14 / N To (2 * 3.14) + (3.14 / 6) Step 2 * 3.14 / N
        If t = True Then
           R = R - 25
           t = False
         Else
           R = R + 25
           t = True
         End If
        X = Xc + R * Cos(Alpha)
        Y = Yc - R * Sin(Alpha)
        Line -(X, Y)
Next Alpha
End Sub

