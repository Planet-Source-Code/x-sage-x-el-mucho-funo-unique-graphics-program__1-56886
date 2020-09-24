VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "El Mucho Funo by x sAGE x"
   ClientHeight    =   4515
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6780
   LinkTopic       =   "Form1"
   ScaleHeight     =   4515
   ScaleWidth      =   6780
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog Common 
      Left            =   3000
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   2400
      Top             =   960
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Height          =   975
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   6720
      TabIndex        =   0
      Top             =   3540
      Width           =   6780
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   255
         Left            =   3000
         TabIndex        =   18
         Top             =   600
         Width           =   855
      End
      Begin MSComctlLib.Slider scrWidth 
         Height          =   255
         Left            =   5280
         TabIndex        =   15
         Top             =   120
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         _Version        =   393216
         Min             =   1
         SelStart        =   1
         Value           =   1
      End
      Begin VB.OptionButton optCircle 
         Caption         =   "Circle"
         Height          =   255
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   600
         Width           =   615
      End
      Begin VB.OptionButton optBox 
         Caption         =   "Box"
         Height          =   255
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   360
         Width           =   615
      End
      Begin VB.OptionButton optLine 
         Caption         =   "Line"
         Height          =   255
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   120
         Value           =   -1  'True
         Width           =   615
      End
      Begin MSComctlLib.Slider Speed 
         Height          =   255
         Left            =   5280
         TabIndex        =   11
         Top             =   360
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         _Version        =   393216
         Min             =   1
         SelStart        =   1
         Value           =   1
      End
      Begin VB.CommandButton cmdUnpause 
         Caption         =   "Unpause"
         Height          =   255
         Left            =   3000
         TabIndex        =   10
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmdPause 
         Caption         =   "Pause"
         Height          =   255
         Left            =   3000
         TabIndex        =   9
         Top             =   120
         Width           =   855
      End
      Begin VB.PictureBox disabb 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   120
         ScaleHeight     =   735
         ScaleWidth      =   2775
         TabIndex        =   1
         Top             =   120
         Width           =   2775
         Begin VB.PictureBox RGBColor 
            Height          =   735
            Left            =   2640
            ScaleHeight     =   675
            ScaleWidth      =   75
            TabIndex        =   8
            Top             =   0
            Width           =   135
         End
         Begin VB.PictureBox bColor 
            Height          =   255
            Left            =   2400
            ScaleHeight     =   195
            ScaleWidth      =   195
            TabIndex        =   7
            Top             =   480
            Width           =   255
         End
         Begin VB.PictureBox gColor 
            Height          =   255
            Left            =   2400
            ScaleHeight     =   195
            ScaleWidth      =   195
            TabIndex        =   6
            Top             =   240
            Width           =   255
         End
         Begin VB.PictureBox rColor 
            Height          =   255
            Left            =   2400
            ScaleHeight     =   195
            ScaleWidth      =   195
            TabIndex        =   5
            Top             =   0
            Width           =   255
         End
         Begin VB.HScrollBar bScroll 
            Height          =   255
            Left            =   0
            Max             =   255
            TabIndex        =   4
            Top             =   480
            Width           =   2415
         End
         Begin VB.HScrollBar gScroll 
            Height          =   255
            Left            =   0
            Max             =   255
            TabIndex        =   3
            Top             =   240
            Width           =   2415
         End
         Begin VB.HScrollBar rScroll 
            Height          =   255
            Left            =   0
            Max             =   255
            TabIndex        =   2
            Top             =   0
            Width           =   2415
         End
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "Width:"
         Height          =   255
         Index           =   1
         Left            =   4680
         TabIndex        =   17
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "Speed:"
         Height          =   255
         Index           =   0
         Left            =   4680
         TabIndex        =   16
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1920
      Top             =   960
   End
   Begin VB.Line Line2 
      X1              =   2400
      X2              =   4920
      Y1              =   2280
      Y2              =   2160
   End
   Begin VB.Line Line1 
      X1              =   360
      X2              =   480
      Y1              =   840
      Y2              =   1800
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim r As Integer
Dim g As Integer
Dim b As Integer
Dim speeed As Integer
Private Sub bScroll_Change()
b = bScroll.Value
End Sub

Private Sub cmdPause_Click()
Timer1.Enabled = False
disabb.Enabled = True
End Sub

Private Sub cmdSave_Click()
On Error Resume Next
Common.Filter = "Windows Bmp (*.bmp)|*.bmp"
Common.ShowSave
SavePicture Me.Image, Common.FileName
End Sub

Private Sub cmdUnpause_Click()
Timer1.Enabled = True
disabb.Enabled = False
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Randomize
Line2.X1 = X
Line2.Y1 = Y
Timer1.Enabled = False
Line2.Visible = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Or Button = 2 Then
Line2.X2 = X
Line2.Y2 = Y
Line2.Visible = True
End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Line2.X2 = X
Line2.Y2 = Y
Timer1.Enabled = True
Line2.Visible = True
End Sub
Private Sub gScroll_Change()
g = gScroll.Value
End Sub

Private Sub rScroll_Change()
r = rScroll.Value
End Sub

Private Sub Speed_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Speed.ToolTipText = Speed.Value
End Sub

Private Sub Timer1_Timer()
speeed = Speed.Value * 10
Do Until speeed = 0
speeed = speeed - 1
On Error Resume Next
b = b + Int(Rnd * 5)
b = b - Int(Rnd * 5)
r = r + Int(Rnd * 5)
r = r - Int(Rnd * 5)
g = g + Int(Rnd * 5)
g = g - Int(Rnd * 5)
If r > 255 Then r = r - 20
If g > 255 Then g = g - 20
If b > 255 Then b = b - 20
rScroll.Value = Abs(r)
gScroll.Value = Abs(g)
bScroll.Value = Abs(b)
r = Abs(r)
g = Abs(g)
b = Abs(b)
If optLine.Value = True Then
Me.Line (Line1.X1, Line1.Y1)-(Line1.X2, Line1.Y2), RGB(Abs(r), Abs(g), Abs(b))
End If
If optBox.Value = True Then
Me.Line (Line1.X1, Line1.Y1)-(Line1.X2, Line1.Y2), RGB(Abs(r), Abs(g), Abs(b)), B
End If
If optCircle.Value = True Then
Me.Circle (Line1.X1, Line1.Y1), (Line1.X2 + Line1.X1 + Line1.Y1 + Line1.Y2) / 50, RGB(r, g, b)
Me.Circle (Line1.X2, Line1.Y2), (Line1.X2 + Line1.X1 + Line1.Y1 + Line1.Y2) / 50, RGB(r, g, b)
End If

If Line1.X1 > Line2.X1 Then
Line1.X1 = Line1.X1 - 1
End If
If Line1.X1 < Line2.X1 Then
Line1.X1 = Line1.X1 + 1
End If
If Line1.Y1 > Line2.Y1 Then
Line1.Y1 = Line1.Y1 - 1
End If
If Line1.Y1 < Line2.Y1 Then
Line1.Y1 = Line1.Y1 + 1
End If
If Line1.X2 > Line2.X2 Then
Line1.X2 = Line1.X2 - 1
End If
If Line1.X2 < Line2.X2 Then
Line1.X2 = Line1.X2 + 1
End If
If Line1.Y2 > Line2.Y2 Then
Line1.Y2 = Line1.Y2 - 1
End If
If Line1.Y2 < Line2.Y2 Then
Line1.Y2 = Line1.Y2 + 1
End If

Loop
End Sub

Private Sub Timer2_Timer()
Me.DrawWidth = scrWidth.Value
rColor.BackColor = Abs(r)
gColor.BackColor = Abs(g)
bColor.BackColor = Abs(b)
RGBColor.BackColor = RGB(Abs(r), Abs(g), Abs(b))
End Sub
