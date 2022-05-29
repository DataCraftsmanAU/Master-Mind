VERSION 5.00
Begin VB.Form Help 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Help"
   ClientHeight    =   9780
   ClientLeft      =   2700
   ClientTop       =   1755
   ClientWidth     =   11250
   Icon            =   "Help.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Help.frx":0E42
   ScaleHeight     =   9780
   ScaleWidth      =   11250
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      DrawMode        =   1  'Blackness
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000003&
      Height          =   495
      Index           =   3
      Left            =   8040
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   9
      ToolTipText     =   "Left Click to cycle through the colours, right click if you want to cycle backwards."
      Top             =   1720
      Width           =   495
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      DrawMode        =   1  'Blackness
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000003&
      Height          =   495
      Index           =   2
      Left            =   8760
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   8
      ToolTipText     =   "Left Click to cycle through the colours, right click if you want to cycle backwards."
      Top             =   1720
      Width           =   495
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      DrawMode        =   1  'Blackness
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000003&
      Height          =   495
      Index           =   1
      Left            =   10200
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   6
      ToolTipText     =   "Left Click to cycle through the colours, right click if you want to cycle backwards."
      Top             =   1720
      Width           =   495
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      DrawMode        =   1  'Blackness
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000003&
      Height          =   495
      Index           =   0
      Left            =   9480
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   4
      ToolTipText     =   "Left Click to cycle through the colours, right click if you want to cycle backwards."
      Top             =   1720
      Width           =   495
   End
   Begin VB.CommandButton Demo 
      Caption         =   "Demo Game"
      Height          =   375
      Left            =   9840
      TabIndex        =   1
      ToolTipText     =   "Click for a preview of a game incase you dont understand."
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   $"Help.frx":99A88
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1815
      Left            =   600
      TabIndex        =   19
      Top             =   7440
      Width           =   4455
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   $"Help.frx":99B66
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1575
      Left            =   600
      TabIndex        =   18
      Top             =   5880
      Width           =   4455
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   $"Help.frx":99C2E
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1815
      Left            =   600
      TabIndex        =   17
      Top             =   4080
      Width           =   4455
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   $"Help.frx":99D2E
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   2535
      Left            =   600
      TabIndex        =   16
      Top             =   1680
      Width           =   4455
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   $"Help.frx":99E5D
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1320
      Left            =   6240
      TabIndex        =   15
      Top             =   8040
      Width           =   4455
   End
   Begin VB.Shape Circ 
      BorderColor     =   &H80000005&
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   7
      Left            =   7560
      Shape           =   3  'Circle
      Top             =   7320
      Width           =   735
   End
   Begin VB.Shape Circ 
      BorderColor     =   &H80000005&
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   6
      Left            =   8400
      Shape           =   3  'Circle
      Top             =   7320
      Width           =   735
   End
   Begin VB.Shape Circ 
      BorderColor     =   &H80000005&
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   5
      Left            =   9240
      Shape           =   3  'Circle
      Top             =   7320
      Width           =   735
   End
   Begin VB.Shape Circ 
      BorderColor     =   &H80000005&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   4
      Left            =   10080
      Shape           =   3  'Circle
      Top             =   7320
      Width           =   735
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "These circles show your attempts at guessing the"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   975
      Left            =   6240
      TabIndex        =   14
      Top             =   7155
      Width           =   1815
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000009&
      BorderStyle     =   3  'Dot
      X1              =   6000
      X2              =   10920
      Y1              =   6960
      Y2              =   6960
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      BorderStyle     =   3  'Dot
      X1              =   6000
      X2              =   10920
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Means that you have got 1  or more of the colours wrong."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   6600
      TabIndex        =   13
      Top             =   6120
      Width           =   4095
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Means that you have got 1 or more of the colours right but in the wrong position."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   6600
      TabIndex        =   12
      Top             =   5400
      Width           =   4335
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Means that you have got 1 or more of the colours right and in the correct position."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   6600
      TabIndex        =   11
      Top             =   4680
      Width           =   4335
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "These are the squares that say how many of the colours you chose matched the code."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   6240
      TabIndex        =   10
      Top             =   3960
      Width           =   4575
   End
   Begin VB.Shape Square 
      BorderColor     =   &H80000005&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   110
      Left            =   6240
      Top             =   4800
      Width           =   255
   End
   Begin VB.Shape Square 
      BorderColor     =   &H80000005&
      FillColor       =   &H000000FF&
      FillStyle       =   7  'Diagonal Cross
      Height          =   255
      Index           =   111
      Left            =   6240
      Top             =   5520
      Width           =   255
   End
   Begin VB.Shape Square 
      BorderColor     =   &H80000005&
      FillColor       =   &H000000FF&
      Height          =   255
      Index           =   113
      Left            =   6240
      Top             =   6240
      Width           =   255
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   $"Help.frx":99F0C
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1215
      Left            =   6200
      TabIndex        =   7
      Top             =   2280
      Width           =   4575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "These Squares are used for choosing your"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   975
      Left            =   6240
      TabIndex        =   5
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Shape Helpback1 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      DrawMode        =   5  'Not Copy Pen
      FillStyle       =   0  'Solid
      Height          =   8055
      Left            =   360
      Top             =   1440
      Width           =   4935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Game Interface"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   855
      Left            =   6960
      TabIndex        =   3
      Top             =   720
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "How To Play"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   855
      Left            =   1320
      TabIndex        =   2
      Top             =   720
      Width           =   2895
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000004&
      BorderStyle     =   2  'Dash
      X1              =   5640
      X2              =   5640
      Y1              =   840
      Y2              =   9600
   End
   Begin VB.Label Mastermind 
      BackStyle       =   0  'Transparent
      Caption         =   "Mastermind"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   855
      Left            =   3960
      TabIndex        =   0
      Top             =   -120
      Width           =   3495
   End
   Begin VB.Shape Helpback2 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      DrawMode        =   5  'Not Copy Pen
      FillStyle       =   0  'Solid
      Height          =   8055
      Left            =   6000
      Top             =   1440
      Width           =   4935
   End
End
Attribute VB_Name = "Help"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Demo_Click()
Help.Hide
Master.Height = 11120
Master.Width = 12000

'========== blanks all the circles =========
For z = 0 To 11
    For x = 0 To 3
        Master.Circ(z & x).FillColor = &H80000007
    Next
Next
'========== /blanks all the circles =========

'========== blanks all the squares =========
For z = 0 To 11
    For x = 0 To 3
        Master.Square(z & x).FillStyle = 1
    Next
Next
'========== /blanks all the squares =========

'1 right place, 1 right colour
Master.Circ(0).FillColor = &HFFFF& 'Yellow
Master.Circ(1).FillColor = &HFFFF& 'Yellow
Master.Circ(2).FillColor = &HFF& 'Red
Master.Circ(3).FillColor = &H80FF& 'Orange
Master.Square(0).FillStyle = 0
Master.Square(1).FillStyle = 7
Master.Square(2).FillStyle = 1
Master.Square(3).FillStyle = 1

'1 right place, 2 right colour
Master.Circ(10).FillColor = &HFF& 'Red
Master.Circ(11).FillColor = &HFFFF& 'Yellow
Master.Circ(12).FillColor = &HFF00& 'Green
Master.Circ(13).FillColor = &H80FF& 'Orange
Master.Square(10).FillStyle = 0
Master.Square(11).FillStyle = 7
Master.Square(12).FillStyle = 7
Master.Square(13).FillStyle = 1

'2 right place, 2 right colour
Master.Circ(20).FillColor = &HFF00& 'Green
Master.Circ(21).FillColor = &HFFFF& 'Yellow
Master.Circ(22).FillColor = &HFF00& 'Green
Master.Circ(23).FillColor = &H80FF& 'Orange
Master.Square(20).FillStyle = 0
Master.Square(21).FillStyle = 0
Master.Square(22).FillStyle = 7
Master.Square(23).FillStyle = 7

'2 right place, 2 right colour
Master.Circ(30).FillColor = &HFFFF& 'Yellow
Master.Circ(31).FillColor = &HFF00& 'Green
Master.Circ(32).FillColor = &HFF00& 'Green
Master.Circ(33).FillColor = &H80FF& 'Orange
Master.Square(30).FillStyle = 0
Master.Square(31).FillStyle = 0
Master.Square(32).FillStyle = 7
Master.Square(33).FillStyle = 7

'win
Master.Circ(40).FillColor = &HFF00&  'Green
Master.Circ(41).FillColor = &HFF00& 'Green
Master.Circ(42).FillColor = &HFFFF& 'Yellow
Master.Circ(43).FillColor = &H80FF& 'Orange
Master.Square(40).FillStyle = 0
Master.Square(41).FillStyle = 0
Master.Square(42).FillStyle = 0
Master.Square(43).FillStyle = 0

MsgBox ("The code was Green, Green, Yellow, Orange.")
MsgBox ("Press OK when you understand a little better.")
Master.Width = 6600
Master.Height = 6045
End Sub

