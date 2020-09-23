VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sine/Cosine Curve Plotter"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5415
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   310
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   361
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Pause"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3960
      TabIndex        =   16
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "C&lose"
      Height          =   375
      Left            =   3960
      TabIndex        =   10
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Controls"
      Height          =   1455
      Left            =   120
      TabIndex        =   1
      Top             =   3120
      Width           =   3555
      Begin VB.OptionButton Option2 
         Caption         =   "Cosine"
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Sine"
         Height          =   495
         Left            =   360
         TabIndex        =   3
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   2640
         MaxLength       =   4
         TabIndex        =   2
         Top             =   300
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "Frequency:"
         Height          =   255
         Left            =   1500
         TabIndex        =   9
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Counter:"
         Height          =   255
         Left            =   1680
         TabIndex        =   8
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Angle:"
         Height          =   255
         Left            =   1800
         TabIndex        =   7
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label3 
         Height          =   255
         Left            =   2640
         TabIndex        =   6
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label2 
         Height          =   255
         Left            =   2640
         TabIndex        =   5
         Top             =   720
         Width           =   375
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Draw"
      Height          =   375
      Left            =   3960
      TabIndex        =   0
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3960
      Top             =   2400
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4560
      Top             =   2400
   End
   Begin VB.Label Label8 
      Caption         =   "270"
      Height          =   255
      Left            =   3900
      TabIndex        =   13
      Top             =   1860
      Width           =   315
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   200
   End
   Begin VB.Line Line4 
      Index           =   3
      X1              =   630
      X2              =   630
      Y1              =   100
      Y2              =   112
   End
   Begin VB.Line Line4 
      Index           =   2
      X1              =   450
      X2              =   450
      Y1              =   100
      Y2              =   112
   End
   Begin VB.Line Line3 
      Index           =   2
      X1              =   540
      X2              =   540
      Y1              =   100
      Y2              =   124
   End
   Begin VB.Line Line3 
      Index           =   1
      X1              =   360
      X2              =   360
      Y1              =   100
      Y2              =   124
   End
   Begin VB.Line Line4 
      Index           =   1
      X1              =   270
      X2              =   270
      Y1              =   100
      Y2              =   112
   End
   Begin VB.Line Line4 
      Index           =   0
      X1              =   90
      X2              =   90
      Y1              =   100
      Y2              =   112
   End
   Begin VB.Line Line3 
      Index           =   0
      X1              =   180
      X2              =   180
      Y1              =   100
      Y2              =   124
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   720
      Y1              =   100
      Y2              =   100
   End
   Begin VB.Label Label1 
      Caption         =   "180"
      Height          =   255
      Left            =   2580
      TabIndex        =   11
      Top             =   1860
      Width           =   375
   End
   Begin VB.Label Label7 
      Caption         =   "90"
      Height          =   255
      Left            =   1260
      TabIndex        =   12
      Top             =   1860
      Width           =   315
   End
   Begin VB.Label Label9 
      Caption         =   "1"
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   14
      Top             =   0
      Width           =   315
   End
   Begin VB.Label Label9 
      Caption         =   "-1"
      Height          =   255
      Index           =   1
      Left            =   60
      TabIndex        =   15
      Top             =   2820
      Width           =   315
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Counter, Sinx, Ypos, Nums, Numc
Dim CCheck As Boolean
Dim X, Y

Sub SinCurve()
Label2.Caption = Counter
Nums = Counter * 0.0174532 'This converts radians into degrees
Sinx = (Sin(Nums) + 1) * 100 'This creates an X coordinate for the line to use
Line -(Ypos, Sinx)
End Sub

Sub CosCurve()
Label2.Caption = Counter
Numc = Counter * 0.0174532
Cosx = (Cos(Numc) - 1) * 100
Line -(Ypos, -Cosx)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyD Then
Command1_Click
ElseIf Command2.Enabled = True And KeyCode = vbKeyP Or KeyCode = vbKeyC Then
Command2_Click
ElseIf KeyCode = vbKeyL Then
Unload Me
End If

End Sub

Private Sub Form_Load()
CCheck = True
Counter = 360
Text1.Text = 1
CurrentY = 100
Option1.Value = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.Caption = X
End Sub

Private Sub Command1_Click()
On Error GoTo DError
Counter = 360 * Text1.Text
Sinx = 0
Ypos = 0
Txt = Text1.Text
CCheck = True
Cls

If Option1.Value = True Then
    CurrentX = 0
    CurrentY = 100
Else
    CurrentX = 0
    CurrentY = 0
End If

Caption = "Drawing Curve..."
Command2.Enabled = True
Command2.Caption = "&Pause"

Timer1.Enabled = True
Timer2.Enabled = True
DError:
Exit Sub
End Sub

Private Sub Command2_Click()

Select Case CCheck
Case True
Timer1.Enabled = False
Timer2.Enabled = False
Command2.Caption = "&Continue"
CCheck = False

Case False
Timer1.Enabled = True
Timer2.Enabled = True
Command2.Caption = "&Pause"
CCheck = True
End Select

End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Text1_Change()

On Error GoTo TError

Counter = 360 * Text1.Text

TError:
Exit Sub
End Sub

Private Sub Text1_Click()
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
Text1.SetFocus
End Sub

Private Sub Timer1_Timer()
Y = Text1.Text * 3 'Sets how far to move line before next calculation
X = 3 'Sets how fast to move curve horizontally
If Text1.Text >= 1 Then
    Counter = Counter - Y
ElseIf Text1.Text < 1 Then
    Counter = Counter + Y
End If

Ypos = Ypos + X 'This makes the curve move horizontally
End Sub

Private Sub Timer2_Timer()

If Option1.Value = True Then
    SinCurve 'Draw a sine curve
Else
    CosCurve 'Draw a cosine curve
End If

If Counter <= 0 Then
    Timer1.Enabled = False 'Stops drawing
    Caption = "Sine/Cosine Curve Plotter"
    Command2.Enabled = False
End If
End Sub


