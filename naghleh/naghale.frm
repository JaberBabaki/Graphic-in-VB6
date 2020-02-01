VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6555
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9180
   LinkTopic       =   "Form1"
   ScaleHeight     =   6555
   ScaleWidth      =   9180
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   375
      Left            =   7920
      TabIndex        =   0
      Top             =   6000
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
ScaleMode = 6
Call cp(45, 50, 40)
End Sub
Private Sub cp(ByVal xc As Integer, ByVal yc As Integer, ByVal r As Integer)
pi = 4 * Atn(1)
Circle (xc, yc), r, vbRed, 0, pi
Circle (xc, yc), r - r / 4, vbRed, 0, pi
Circle (xc, yc), r - r / 1.1, vbRed, 0, pi
Line (xc - r, yc + 10)-(xc + r, yc + 10), vbRed
Line (xc - r, yc + 10)-(xc - r, yc), vbRed
Line (xc + r, yc + 10)-(xc + r, yc), vbRed
Line (xc - (r - r / 4), yc)-(xc - (r - r / 1.1), yc), vbRed
Line (xc + (r - r / 1.1), yc)-(xc + (r - r / 4), yc), vbRed
PSet (xc, yc), vbRed
CurrentX = xc - r
CurrentY = yc + 10
For i = 0 To 2 * r - 2
CurrentX = (xc - r) + (i + 1)
 Line -Step(0, -1)
CurrentY = CurrentY + 1

If i Mod 10 = 0 Then
Line -Step(0, -2)
CurrentY = CurrentY - 5
CurrentX = CurrentX - 1
Print i / 10;
CurrentY = CurrentY + 7
 CurrentX = (xc - r) + (i + 1)
End If

If i Mod 5 = 0 Then
Line -Step(0, -1.5), vbBlue
CurrentY = CurrentY + 1.5
End If
Next

s = 180: m = 1
For i = pi To 2 * pi Step pi / 178
  x = xc + r * Cos(i)
  y = yc + r * Sin(i)
  xm = xc + r / 1.05 * Cos(i)
  ym = yc + r / 1.05 * Sin(i)
  xp = xc + r / 1.1 * Cos(i)
  yp = yc + r / 1.1 * Sin(i)
  PSet (x, y)
  Line (x, y)-(xm, ym), vbRed
  If (s Mod 10 = 0) Or (s = 1) Then
  Line (x, y)-(xp, yp), vbBlue
 If s < 90 Then CurrentX = CurrentX - 3
  
  Print s;
  

  End If
   s = s - 1
 Next


End Sub
