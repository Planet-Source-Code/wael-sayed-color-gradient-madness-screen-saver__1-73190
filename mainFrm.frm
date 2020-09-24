VERSION 5.00
Begin VB.Form mainFrm 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   11160
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13440
   DrawWidth       =   3
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   744
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   896
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picColors 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   122
      Left            =   6960
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   24
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picColors 
      BackColor       =   &H00808080&
      Height          =   255
      Index           =   121
      Left            =   6720
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   23
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picColors 
      BackColor       =   &H00808080&
      Height          =   255
      Index           =   99
      Left            =   1440
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   22
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picColors 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   98
      Left            =   1200
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   21
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picColors 
      BackColor       =   &H0004A7FF&
      Height          =   255
      Index           =   120
      Left            =   6480
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   20
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picColors 
      BackColor       =   &H0004FF8E&
      Height          =   255
      Index           =   118
      Left            =   6000
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   19
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picColors 
      BackColor       =   &H0004A7FF&
      Height          =   255
      Index           =   100
      Left            =   1680
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   18
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picColors 
      BackColor       =   &H0000FFFF&
      Height          =   255
      Index           =   101
      Left            =   1920
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   17
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picColors 
      BackColor       =   &H0004FF8E&
      Height          =   255
      Index           =   102
      Left            =   2160
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   16
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picColors 
      BackColor       =   &H0000C000&
      Height          =   255
      Index           =   103
      Left            =   2400
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   15
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picColors 
      BackColor       =   &H00404000&
      Height          =   255
      Index           =   104
      Left            =   2640
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   14
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picColors 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   105
      Left            =   2880
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   13
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picColors 
      BackColor       =   &H00C0C000&
      Height          =   255
      Index           =   106
      Left            =   3120
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   12
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picColors 
      BackColor       =   &H00FFFF80&
      Height          =   255
      Index           =   107
      Left            =   3360
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   11
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picColors 
      BackColor       =   &H00C0E0FF&
      Height          =   255
      Index           =   108
      Left            =   3600
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   10
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picColors 
      BackColor       =   &H00C0C0FF&
      Height          =   255
      Index           =   109
      Left            =   3840
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   9
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picColors 
      BackColor       =   &H000000FF&
      Height          =   255
      Index           =   110
      Left            =   4080
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   8
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   9480
      Top             =   120
   End
   Begin VB.PictureBox picColors 
      BackColor       =   &H0000FFFF&
      Height          =   255
      Index           =   119
      Left            =   6240
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   7
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picColors 
      BackColor       =   &H0000C000&
      Height          =   255
      Index           =   117
      Left            =   5760
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   6
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picColors 
      BackColor       =   &H00404000&
      Height          =   255
      Index           =   116
      Left            =   5520
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   5
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picColors 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   115
      Left            =   5280
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picColors 
      BackColor       =   &H00C0C000&
      Height          =   255
      Index           =   114
      Left            =   5040
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picColors 
      BackColor       =   &H00FFFF80&
      Height          =   255
      Index           =   113
      Left            =   4800
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picColors 
      BackColor       =   &H00C0E0FF&
      Height          =   255
      Index           =   112
      Left            =   4560
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picColors 
      BackColor       =   &H00C0C0FF&
      Height          =   255
      Index           =   111
      Left            =   4320
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "mainFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Const SizeX = 24
Private Type RGBset
    Angle As Integer
    R(0 To SizeX)
    G(0 To SizeX)
    B(0 To SizeX)
    Count As Integer
End Type
Dim gradtemp As RGBset


Dim Centers() As Single, ColorsOrder(0 To SizeX / 2) As Double, DelayFlag As Boolean, R As Integer
Private Declare Function GetTickCount Lib "kernel32" () As Long 'For timer
Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function ShowCursor Lib "user32" (ByVal fShow As Integer) As Integer



Public Sub Delay(Period As Integer)




Dim dtStart As Long
Dim dtEnd As Long
Dim result As Long
Dim i As Integer


If DelayFlag Then Exit Sub
dtStart = GetTickCount

eee:

DoEvents

dtEnd = GetTickCount

If dtEnd - dtStart < Period Then GoTo eee




End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

Unload Me

End Sub

Private Sub Form_Load()


On Error GoTo kkk

10  If App.PrevInstance Then End
30  sCommand = LCase$(Trim$(Command()))
40  If sCommand <> "/s" Then End
50   Call SetWindowPos(Me.hwnd, -1, 0&, 0&, Me.ScaleWidth, Me.ScaleHeight, &H40)
60 hhh:
70  If ShowCursor(False) >= 0 Then GoTo hhh
80
85  Me.WindowState = 2
90  Me.Scale (-Me.ScaleWidth / 2, Me.ScaleHeight / 2)-(Me.ScaleWidth / 2, -Me.ScaleHeight / 2)
95  Me.Show
100  gradtemp.Count = picColors.Count
110  DelayFlag = True
120  ReDim Centers(1 To 1, 1 To 2)

150  For X% = 0 To UBound(ColorsOrder)
160     ColorsOrder(X) = picColors(picColors.LBound + X).BackColor
170  Next X

Randomize
R% = (Rnd * (Abs(Me.ScaleHeight / 8))) + Abs(Me.ScaleHeight / 8)
130  Centers(1, 1) = Rnd * R / 2
140  Centers(1, 2) = Rnd * R / 2
180  ColorShift
190  DrawGrad Me, gradtemp
200  Exit Sub


kkk: Me.WindowState = 1
MsgBox "Form Load: " + CStr(Erl) + "   " + CStr(Me.ScaleHeight) + "     " + CStr(Me.ScaleWidth) + "  " + Command()
Unload Me


End Sub





Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Unload Me


End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)


Static Z As Integer

Z = Z + 1

If Z > 25 Then Unload Me

End Sub

Private Sub Form_Unload(Cancel As Integer)


hhh:
If ShowCursor(True) < 0 Then GoTo hhh

End


End Sub

Private Function Red(ByVal Color As Long) As Integer
    Red = Color Mod &H100
End Function
Private Function Green(ByVal Color As Long) As Integer
    Green = (Color \ &H100) Mod &H100
End Function
Private Function Blue(ByVal Color As Long) As Integer
    Blue = (Color \ &H10000) Mod &H100
End Function

Private Sub DrawGrad(obj As Object, rgbs As RGBset)


On Error GoTo fff

Dim Colors As Integer
Static rad_ang As Single, Clear As Integer



Clear = Clear + 1


obj.DrawMode = 13
10  C = Rnd
40  If C < 1 And C > 0.95 Then
50     obj.DrawMode = 4
80  End If

90   C = Rnd
100  If C < 0.3 And C > 0 Then
110     DelayFlag = True
120  ElseIf C < 1 And C > 0.7 Then
130     DelayFlag = False
140  End If
150


170  If Clear > Int(Rnd * 50) + 100 Or Clear > 1000 Then
        Clear = 0
180     R% = (Abs(obj.ScaleWidth)) * 1.3
        If Rnd > 0.3 Then obj.DrawMode = 1
        If Rnd > 0.5 Then Centers(1, 1) = 0
        If Rnd > 0.7 Then Centers(1, 2) = 0
        GoTo ooo
190  Else
200     R% = (Rnd * (Abs(obj.ScaleHeight / 8))) + Abs(obj.ScaleHeight / 8)
210  End If

470      KM! = Rnd * 0.2
480      KL! = Rnd * 0.05
ooo:

270      r2 = rgbs.R(0)
280      g2 = rgbs.G(0)
290      b2 = rgbs.B(0)
330      Colors = Rnd * (rgbs.Count / 2 - 2) + 2
340      h = Int(2 * 22 / 7 * R / Colors / 2)
350      Facto = h * Colors * 2
360      r2 = rgbs.R((UBound(rgbs.R) / 2) - Colors)
370      g2 = rgbs.G((UBound(rgbs.R) / 2) - Colors)
380      b2 = rgbs.B((UBound(rgbs.R) / 2) - Colors)
390      rad_ang = rad_ang + ((-1) ^ Int(Rnd * 3)) + 0.01
400      If rad_ang > 100 Then rad_ang = 0
440      Z = (-1) ^ (Int(Rnd * 2) + 1)
450      Num = Int(Rnd * 20)
460      Num1 = Int(Rnd * 5)
490      For C = (UBound(rgbs.R) / 2) - Colors To (UBound(rgbs.R) / 2) + Colors
500          rd = (rgbs.R(C) - r2) / h
510          gd = (rgbs.G(C) - g2) / h
520          bd = (rgbs.B(C) - b2) / h
530          For Y = 0 To (h - 1)
540              r2 = r2 + rd
550              g2 = g2 + gd
560              b2 = b2 + bd
570              rad_ang = rad_ang + Z * (44 / 7) / Facto
580              For q% = 1 To 4
590                  If q = 1 Then i% = 1: J% = 1
600                  If q = 2 Then i% = -1: J% = 1
610                  If q = 3 Then i% = 1: J% = -1
620                  If q = 4 Then i% = -1: J% = -1
630                  X1 = i * Centers(1, 1)
640                  Y1 = J * Centers(1, 2)
650                  X2 = R * (0.9 - KM - KL + KM * Cos(Num * rad_ang) + KL * Cos(Num * Num1 * rad_ang)) * Cos(rad_ang) + Centers(1, 1)
660                  Y2 = R * (0.9 - KM - KL + KM * Cos(Num * rad_ang) + KL * Cos(Num * Num1 * rad_ang)) * Sin(rad_ang) + Centers(1, 2)
670                  obj.Line (X1, Y1)-(i * X2, J * Y2), RGB(r2, g2, b2)
680              Next
690          Next Y
700          Delay Rnd * 20 + 2
710      Next C


s% = Int(Rnd * 2)



730  Centers(1, 1) = X2
740  Centers(1, 2) = Y2
750  Timer1.Enabled = True


Exit Sub

fff: Me.WindowState = 1
MsgBox "Drawing: " + CStr(Erl)
Unload Me

End Sub

Private Sub Timer1_Timer()
        
Static m As Single, d As Single, n As Integer, w As Integer
        
Timer1.Enabled = False
ColorShift

If Not (Centers(1, 1) > -Abs(Me.ScaleWidth) / 2 And Centers(1, 1) < Abs(Me.ScaleWidth) / 2) Or _
   Not (Centers(1, 2) > -Abs(Me.ScaleHeight) / 2 And Centers(1, 2) < Abs(Me.ScaleHeight) / 2) Then
       Centers(1, 1) = Rnd * Abs(Me.ScaleHeight) / 2
       Centers(1, 2) = Rnd * Abs(Me.ScaleHeight) / 2
End If


Delay Rnd * 700 + 10
DrawGrad Me, gradtemp


End Sub



Public Sub ColorScramble(ReOrder As Boolean)

On Error GoTo zzz
'Randomize

10  If ReOrder = False Then
20
30      For X% = picColors.LBound To picColors.LBound + picColors.Count / 2 ' - 2
40         A = picColors.LBound + Int(Rnd * (picColors.Count / 2))
50         B = picColors.LBound + Int(Rnd * (picColors.Count / 2))
60         C = picColors(A).BackColor
70         picColors(A).BackColor = picColors(B).BackColor
80         picColors(B).BackColor = C
90      Next
100  Else
110      For X% = picColors.LBound To picColors.LBound + picColors.Count / 2 ' - 2
120          picColors(X).BackColor = ColorsOrder(X - picColors.LBound)
130      Next X
         Z = Int(Rnd * 4) + 3
         For X% = 1 To Z ' Int(Rnd * 4) + 3
             ColorShift True
         Next
140  End If
150  Exit Sub


zzz: Me.WindowState = 1
MsgBox "ColorScramble: " + CStr(Erl)
Unload Me


'For x% = picColors.LBound + picColors.Count / 2 To picColors.ubound
'    picColors(x).BackColor = picColors(picColors.LBound + picColors.ubound - x).BackColor
'Next x
'For i = 0 To picColors.Count - 1
'    gradtemp.R(i) = Red(picColors(i + picColors.LBound).BackColor)
'    gradtemp.G(i) = Green(picColors(i + picColors.LBound).BackColor)
'    gradtemp.B(i) = Blue(picColors(i + picColors.LBound).BackColor)
'Next i




End Sub
Public Sub ColorShift(Optional Without_Scramble As Boolean)



On Error GoTo eee

If Without_Scramble Then GoTo aaa

'10  Randomize
20  C = Rnd
30  If C < 0.5 And C > 0 Then
40     ColorScramble (True)
50  ElseIf C < 1 And C > 0.95 Then
60     ColorScramble (False)
70  End If

aaa:
80  C = picColors(picColors.LBound).BackColor
90  For X% = picColors.LBound To picColors.LBound + picColors.Count / 2 '- 1
100      picColors(X).BackColor = picColors(X + 1).BackColor
110  Next
120  picColors(X - 1).BackColor = C
130  For X% = picColors.LBound + picColors.Count / 2 To picColors.UBound
140      picColors(X).BackColor = picColors(picColors.LBound + picColors.UBound - X).BackColor
150  Next X
160  For i = 0 To picColors.Count - 1
170      gradtemp.R(i) = Red(picColors(i + picColors.LBound).BackColor)
180      gradtemp.G(i) = Green(picColors(i + picColors.LBound).BackColor)
190      gradtemp.B(i) = Blue(picColors(i + picColors.LBound).BackColor)
200  Next i


Exit Sub
eee: Me.WindowState = 1
MsgBox "ColorShift: " + CStr(Erl)
Unload Me

End Sub

