VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Sharif Aly - www.Planet-Source-code.com"
   ClientHeight    =   3045
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5100
   LinkTopic       =   "Form1"
   ScaleHeight     =   3045
   ScaleWidth      =   5100
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Enter"
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Text            =   "# Number"
      Top             =   720
      Width           =   975
   End
   Begin VB.HScrollBar HScrollx 
      Height          =   255
      LargeChange     =   5
      Left            =   240
      Max             =   100
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   2160
      Value           =   100
      Width           =   4095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "PSCODE - Updated By SafSoft"
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   2640
      Width           =   3855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " 255"
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Transparent transparent form With Grade"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
    Call Rainbow
End Sub
Private Sub Rainbow()
    On Error Resume Next
    Dim Position As Integer, Red As Integer, Green As _
    Integer, Blue As Integer
    Dim ScaleFactor As Double, Length As Integer
    ScaleFactor = Me.ScaleWidth / (255 * 6)
    Length = Int(ScaleFactor * 255)
    Position = 0
    Red = 255
    Blue = 1
    For Green = 1 To Length
        Me.Line (Position, 0)-(Position, Me.ScaleHeight), _
        RGB(Red, Green \ ScaleFactor, Blue)
        Position = Position + 1
    Next Green
    For Red = Length To 1 Step -1
        Me.Line (Position, 0)-(Position, Me.ScaleHeight), _
        RGB(Red \ ScaleFactor, Green, Blue)
        Position = Position + 1
    Next Red
    For Blue = 0 To Length
        Me.Line (Position, 0)-(Position, Me.ScaleHeight), _
        RGB(Red, Green, Blue \ ScaleFactor)
        Position = Position + 1
    Next Blue
    For Green = Length To 1 Step -1
        Me.Line (Position, 0)-(Position, Me.ScaleHeight), _
        RGB(Red, Green \ ScaleFactor, Blue)
        Position = Position + 1
    Next Green
    For Red = 1 To Length
        Me.Line (Position, 0)-(Position, Me.ScaleHeight), _
        RGB(Red \ ScaleFactor, Green, Blue)
        Position = Position + 1
    Next Red
    For Blue = Length To 1 Step -1
        Me.Line (Position, 0)-(Position, Me.ScaleHeight), _
        RGB(Red, Green, Blue \ ScaleFactor)
        Position = Position + 1
    Next Blue
End Sub

Private Sub Form_Load()
MakeTransparent Me.hWnd, Label2.Caption
    Me.AutoRedraw = True
    Me.ScaleMode = vbTwips
Label2.Caption = HScrollx.Value * 1
End Sub

Private Sub HScrollx_Change()
MakeTransparent Me.hWnd, Label2.Caption
Label2.Caption = HScrollx.Value * 1
Label2.Caption = HScrollx.Value
End Sub

Private Sub HScrollx_Scroll()
Label2.Caption = HScrollx.Value * 1
MakeTransparent Me.hWnd, Label2.Caption
End Sub

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text1.Text = Clear
End Sub
