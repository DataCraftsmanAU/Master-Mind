VERSION 5.00
Begin VB.Form Master 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mastermind"
   ClientHeight    =   5250
   ClientLeft      =   2415
   ClientTop       =   1455
   ClientWidth     =   6540
   FillColor       =   &H80000003&
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Main.frx":0E42
   ScaleHeight     =   5250
   ScaleWidth      =   6540
   Begin VB.CommandButton m3 
      Caption         =   "Hard"
      Height          =   495
      Left            =   3600
      TabIndex        =   12
      ToolTipText     =   "Start a 2 Player Game, with Hard Difficulty."
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton m2 
      Caption         =   "Normal"
      Height          =   495
      Left            =   3600
      TabIndex        =   11
      ToolTipText     =   "Start a 2 Player Game, with Normal Difficulty."
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton m1 
      Caption         =   "Easy"
      Height          =   495
      Left            =   3600
      TabIndex        =   10
      ToolTipText     =   "Start a 2 Player Game, with Easy Difficulty."
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton s3 
      Caption         =   "Hard"
      Height          =   495
      Left            =   1680
      TabIndex        =   9
      ToolTipText     =   "Start a Single Player Game, with Hard Difficulty."
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton s2 
      Caption         =   "Normal"
      Height          =   495
      Left            =   1680
      TabIndex        =   8
      ToolTipText     =   "Start a Single Player Game, with Normal Difficulty."
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton s1 
      Caption         =   "Easy"
      Height          =   495
      Left            =   1680
      TabIndex        =   7
      ToolTipText     =   "Start a Single Player Game, with Easy Difficulty."
      Top             =   2040
      Width           =   1095
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      DrawMode        =   1  'Blackness
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000005&
      Height          =   495
      Index           =   3
      Left            =   9480
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   6
      ToolTipText     =   "Left Click to cycle through the colours, right click if you want to cycle backwards."
      Top             =   9480
      Width           =   495
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      DrawMode        =   1  'Blackness
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000005&
      Height          =   495
      Index           =   2
      Left            =   8640
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   5
      ToolTipText     =   "Left Click to cycle through the colours, right click if you want to cycle backwards."
      Top             =   9480
      Width           =   495
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      DrawMode        =   1  'Blackness
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000005&
      Height          =   495
      Index           =   1
      Left            =   7800
      ScaleHeight     =   2.063
      ScaleMode       =   0  'User
      ScaleWidth      =   4.125
      TabIndex        =   4
      ToolTipText     =   "Left Click to cycle through the colours, right click if you want to cycle backwards."
      Top             =   9480
      Width           =   495
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      DrawMode        =   1  'Blackness
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000003&
      Height          =   495
      Index           =   0
      Left            =   6960
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   3
      ToolTipText     =   "Left Click to cycle through the colours, right click if you want to cycle backwards."
      Top             =   9480
      Width           =   495
   End
   Begin VB.CommandButton Pick 
      Caption         =   "Pick Colours!"
      Height          =   735
      Left            =   10440
      TabIndex        =   2
      ToolTipText     =   "Click me to either start a new game, choose/generate a code to try and crack, or place the 4 colours onto the game field."
      Top             =   9360
      Width           =   1215
   End
   Begin VB.CommandButton exitbut 
      BackColor       =   &H00000000&
      Caption         =   "Exit"
      Height          =   495
      Left            =   2640
      MaskColor       =   &H00000000&
      TabIndex        =   1
      ToolTipText     =   "Don't quit now."
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton helpbut 
      BackColor       =   &H80000007&
      Caption         =   "Help"
      Height          =   495
      Left            =   2640
      TabIndex        =   0
      ToolTipText     =   "Click me if you didn't read the user documentation."
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label gamenum 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   495
      Left            =   2880
      TabIndex        =   29
      Top             =   8280
      Width           =   735
   End
   Begin VB.Label labgame 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Games: "
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   600
      TabIndex        =   28
      ToolTipText     =   "Displays how many games have been played."
      Top             =   8280
      Width           =   2175
   End
   Begin VB.Label msg 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   360
      TabIndex        =   27
      Top             =   9480
      Width           =   5895
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00808080&
      BorderStyle     =   3  'Dot
      X1              =   10440
      X2              =   11640
      Y1              =   8340
      Y2              =   8340
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00808080&
      BorderStyle     =   3  'Dot
      X1              =   10440
      X2              =   11640
      Y1              =   7620
      Y2              =   7620
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00808080&
      BorderStyle     =   3  'Dot
      X1              =   10440
      X2              =   11640
      Y1              =   6900
      Y2              =   6900
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00808080&
      BorderStyle     =   3  'Dot
      X1              =   10440
      X2              =   11640
      Y1              =   6180
      Y2              =   6180
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00808080&
      BorderStyle     =   3  'Dot
      X1              =   10440
      X2              =   11640
      Y1              =   5460
      Y2              =   5460
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00808080&
      BorderStyle     =   3  'Dot
      X1              =   10440
      X2              =   11640
      Y1              =   4740
      Y2              =   4740
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00808080&
      BorderStyle     =   3  'Dot
      X1              =   10440
      X2              =   11640
      Y1              =   4020
      Y2              =   4020
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00808080&
      BorderStyle     =   3  'Dot
      X1              =   10440
      X2              =   11640
      Y1              =   3300
      Y2              =   3300
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      BorderStyle     =   3  'Dot
      X1              =   10440
      X2              =   11640
      Y1              =   2580
      Y2              =   2580
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      BorderStyle     =   3  'Dot
      X1              =   10440
      X2              =   11640
      Y1              =   1860
      Y2              =   1860
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   3  'Dot
      X1              =   10440
      X2              =   11640
      Y1              =   1140
      Y2              =   1140
   End
   Begin VB.Label bestnum 
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   495
      Left            =   3120
      TabIndex        =   26
      Top             =   8880
      Width           =   3135
   End
   Begin VB.Label labturn2 
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   495
      Left            =   3720
      TabIndex        =   25
      Top             =   7560
      Width           =   1455
   End
   Begin VB.Label labturn1 
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   495
      Left            =   1800
      TabIndex        =   24
      Top             =   7560
      Width           =   1335
   End
   Begin VB.Label labwin1 
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   495
      Left            =   1800
      TabIndex        =   23
      Top             =   6840
      Width           =   1455
   End
   Begin VB.Label labwin2 
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   495
      Left            =   3720
      TabIndex        =   22
      Top             =   6840
      Width           =   1455
   End
   Begin VB.Label labbest 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "Best Game:"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   1080
      TabIndex        =   21
      ToolTipText     =   "Who holds the best game so far, and how many turns was it in."
      Top             =   8880
      Width           =   1935
   End
   Begin VB.Label labwin 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "Wins:"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   600
      TabIndex        =   20
      ToolTipText     =   "How many wins you have."
      Top             =   6840
      Width           =   1095
   End
   Begin VB.Label labguess 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Scoreboard "
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   480
      TabIndex        =   19
      ToolTipText     =   "Says whos turn it is, or just says scoreboard, whatever floats your boat."
      Top             =   5520
      Width           =   5655
   End
   Begin VB.Label labturn 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "Turns:"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   480
      TabIndex        =   18
      ToolTipText     =   "How many turns you've used."
      Top             =   7560
      Width           =   1215
   End
   Begin VB.Label labname2 
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "Player 2"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   615
      Left            =   3720
      TabIndex        =   17
      ToolTipText     =   "Player 2's Name"
      Top             =   6240
      Width           =   2295
   End
   Begin VB.Label labname1 
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "Player 1"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   615
      Left            =   1560
      TabIndex        =   16
      ToolTipText     =   "Player 1's Name"
      Top             =   6240
      Width           =   2055
   End
   Begin VB.Label mlab 
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "2 Player"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   3480
      TabIndex        =   15
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label slab 
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "Single Player"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   1080
      TabIndex        =   14
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label mmlab 
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "Mastermind"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   68.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1575
      Left            =   120
      TabIndex        =   13
      ToolTipText     =   "Mastermind :D"
      Top             =   -240
      Width           =   6255
   End
   Begin VB.Shape Circ 
      BorderColor     =   &H80000005&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   113
      Left            =   9360
      Shape           =   3  'Circle
      Top             =   8400
      Width           =   735
   End
   Begin VB.Shape Circ 
      BorderColor     =   &H80000005&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   112
      Left            =   8520
      Shape           =   3  'Circle
      Top             =   8400
      Width           =   735
   End
   Begin VB.Shape Circ 
      BorderColor     =   &H80000005&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   111
      Left            =   7680
      Shape           =   3  'Circle
      Top             =   8400
      Width           =   735
   End
   Begin VB.Shape Circ 
      BorderColor     =   &H80000005&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   110
      Left            =   6840
      Shape           =   3  'Circle
      Top             =   8400
      Width           =   735
   End
   Begin VB.Shape Circ 
      BorderColor     =   &H80000005&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   103
      Left            =   9360
      Shape           =   3  'Circle
      Top             =   7680
      Width           =   735
   End
   Begin VB.Shape Circ 
      BorderColor     =   &H80000005&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   102
      Left            =   8520
      Shape           =   3  'Circle
      Top             =   7680
      Width           =   735
   End
   Begin VB.Shape Circ 
      BorderColor     =   &H80000005&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   101
      Left            =   7680
      Shape           =   3  'Circle
      Top             =   7680
      Width           =   735
   End
   Begin VB.Shape Circ 
      BorderColor     =   &H80000005&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   100
      Left            =   6840
      Shape           =   3  'Circle
      Top             =   7680
      Width           =   735
   End
   Begin VB.Shape Circ 
      BorderColor     =   &H80000005&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   93
      Left            =   9360
      Shape           =   3  'Circle
      Top             =   6960
      Width           =   735
   End
   Begin VB.Shape Circ 
      BorderColor     =   &H80000005&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   92
      Left            =   8520
      Shape           =   3  'Circle
      Top             =   6960
      Width           =   735
   End
   Begin VB.Shape Circ 
      BorderColor     =   &H80000005&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   91
      Left            =   7680
      Shape           =   3  'Circle
      Top             =   6960
      Width           =   735
   End
   Begin VB.Shape Circ 
      BorderColor     =   &H80000005&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   90
      Left            =   6840
      Shape           =   3  'Circle
      Top             =   6960
      Width           =   735
   End
   Begin VB.Shape Circ 
      BorderColor     =   &H80000005&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   83
      Left            =   9360
      Shape           =   3  'Circle
      Top             =   6240
      Width           =   735
   End
   Begin VB.Shape Circ 
      BorderColor     =   &H80000005&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   82
      Left            =   8520
      Shape           =   3  'Circle
      Top             =   6240
      Width           =   735
   End
   Begin VB.Shape Circ 
      BorderColor     =   &H80000005&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   81
      Left            =   7680
      Shape           =   3  'Circle
      Top             =   6240
      Width           =   735
   End
   Begin VB.Shape Circ 
      BorderColor     =   &H80000005&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   80
      Left            =   6840
      Shape           =   3  'Circle
      Top             =   6240
      Width           =   735
   End
   Begin VB.Shape Circ 
      BorderColor     =   &H80000005&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   73
      Left            =   9360
      Shape           =   3  'Circle
      Top             =   5520
      Width           =   735
   End
   Begin VB.Shape Circ 
      BorderColor     =   &H80000005&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   72
      Left            =   8520
      Shape           =   3  'Circle
      Top             =   5520
      Width           =   735
   End
   Begin VB.Shape Circ 
      BorderColor     =   &H80000005&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   71
      Left            =   7680
      Shape           =   3  'Circle
      Top             =   5520
      Width           =   735
   End
   Begin VB.Shape Circ 
      BorderColor     =   &H80000005&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   70
      Left            =   6840
      Shape           =   3  'Circle
      Top             =   5520
      Width           =   735
   End
   Begin VB.Shape Circ 
      BorderColor     =   &H80000005&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   63
      Left            =   9360
      Shape           =   3  'Circle
      Top             =   4800
      Width           =   735
   End
   Begin VB.Shape Circ 
      BorderColor     =   &H80000005&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   62
      Left            =   8520
      Shape           =   3  'Circle
      Top             =   4800
      Width           =   735
   End
   Begin VB.Shape Circ 
      BorderColor     =   &H80000005&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   61
      Left            =   7680
      Shape           =   3  'Circle
      Top             =   4800
      Width           =   735
   End
   Begin VB.Shape Circ 
      BorderColor     =   &H80000005&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   60
      Left            =   6840
      Shape           =   3  'Circle
      Top             =   4800
      Width           =   735
   End
   Begin VB.Shape Circ 
      BorderColor     =   &H80000005&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   53
      Left            =   9360
      Shape           =   3  'Circle
      Top             =   4080
      Width           =   735
   End
   Begin VB.Shape Circ 
      BorderColor     =   &H80000005&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   52
      Left            =   8520
      Shape           =   3  'Circle
      Top             =   4080
      Width           =   735
   End
   Begin VB.Shape Circ 
      BorderColor     =   &H80000005&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   51
      Left            =   7680
      Shape           =   3  'Circle
      Top             =   4080
      Width           =   735
   End
   Begin VB.Shape Circ 
      BorderColor     =   &H80000005&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   50
      Left            =   6840
      Shape           =   3  'Circle
      Top             =   4080
      Width           =   735
   End
   Begin VB.Shape Circ 
      BorderColor     =   &H80000005&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   43
      Left            =   9360
      Shape           =   3  'Circle
      Top             =   3360
      Width           =   735
   End
   Begin VB.Shape Circ 
      BorderColor     =   &H80000005&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   42
      Left            =   8520
      Shape           =   3  'Circle
      Top             =   3360
      Width           =   735
   End
   Begin VB.Shape Circ 
      BorderColor     =   &H80000005&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   41
      Left            =   7680
      Shape           =   3  'Circle
      Top             =   3360
      Width           =   735
   End
   Begin VB.Shape Circ 
      BorderColor     =   &H80000005&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   40
      Left            =   6840
      Shape           =   3  'Circle
      Top             =   3360
      Width           =   735
   End
   Begin VB.Shape Circ 
      BorderColor     =   &H80000005&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   33
      Left            =   9360
      Shape           =   3  'Circle
      Top             =   2640
      Width           =   735
   End
   Begin VB.Shape Circ 
      BorderColor     =   &H80000005&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   32
      Left            =   8520
      Shape           =   3  'Circle
      Top             =   2640
      Width           =   735
   End
   Begin VB.Shape Circ 
      BorderColor     =   &H80000005&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   31
      Left            =   7680
      Shape           =   3  'Circle
      Top             =   2640
      Width           =   735
   End
   Begin VB.Shape Circ 
      BorderColor     =   &H80000005&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   30
      Left            =   6840
      Shape           =   3  'Circle
      Top             =   2640
      Width           =   735
   End
   Begin VB.Shape Circ 
      BorderColor     =   &H80000005&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   23
      Left            =   9360
      Shape           =   3  'Circle
      Top             =   1920
      Width           =   735
   End
   Begin VB.Shape Circ 
      BorderColor     =   &H80000005&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   22
      Left            =   8520
      Shape           =   3  'Circle
      Top             =   1920
      Width           =   735
   End
   Begin VB.Shape Circ 
      BorderColor     =   &H80000005&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   21
      Left            =   7680
      Shape           =   3  'Circle
      Top             =   1920
      Width           =   735
   End
   Begin VB.Shape Circ 
      BorderColor     =   &H80000005&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   20
      Left            =   6840
      Shape           =   3  'Circle
      Top             =   1920
      Width           =   735
   End
   Begin VB.Shape Circ 
      BorderColor     =   &H80000005&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   13
      Left            =   9360
      Shape           =   3  'Circle
      Top             =   1200
      Width           =   735
   End
   Begin VB.Shape Circ 
      BorderColor     =   &H80000005&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   12
      Left            =   8520
      Shape           =   3  'Circle
      Top             =   1200
      Width           =   735
   End
   Begin VB.Shape Circ 
      BorderColor     =   &H80000005&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   11
      Left            =   7680
      Shape           =   3  'Circle
      Top             =   1200
      Width           =   735
   End
   Begin VB.Shape Circ 
      BorderColor     =   &H80000005&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   10
      Left            =   6840
      Shape           =   3  'Circle
      Top             =   1200
      Width           =   735
   End
   Begin VB.Shape Circ 
      BorderColor     =   &H80000005&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   3
      Left            =   9360
      Shape           =   3  'Circle
      Top             =   480
      Width           =   735
   End
   Begin VB.Shape Circ 
      BorderColor     =   &H80000005&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   2
      Left            =   8520
      Shape           =   3  'Circle
      Top             =   480
      Width           =   735
   End
   Begin VB.Shape Circ 
      BorderColor     =   &H80000005&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   1
      Left            =   7680
      Shape           =   3  'Circle
      Top             =   480
      Width           =   735
   End
   Begin VB.Shape Square 
      BorderColor     =   &H80000005&
      FillColor       =   &H000000FF&
      Height          =   255
      Index           =   113
      Left            =   11160
      Top             =   8760
      Width           =   255
   End
   Begin VB.Shape Square 
      BorderColor     =   &H80000005&
      FillColor       =   &H000000FF&
      Height          =   255
      Index           =   112
      Left            =   10680
      Top             =   8760
      Width           =   255
   End
   Begin VB.Shape Square 
      BorderColor     =   &H80000005&
      FillColor       =   &H000000FF&
      Height          =   255
      Index           =   111
      Left            =   11160
      Top             =   8400
      Width           =   255
   End
   Begin VB.Shape Square 
      BorderColor     =   &H80000005&
      FillColor       =   &H000000FF&
      Height          =   255
      Index           =   110
      Left            =   10680
      Top             =   8400
      Width           =   255
   End
   Begin VB.Shape Square 
      BorderColor     =   &H80000005&
      FillColor       =   &H000000FF&
      Height          =   255
      Index           =   103
      Left            =   11160
      Top             =   8040
      Width           =   255
   End
   Begin VB.Shape Square 
      BorderColor     =   &H80000005&
      FillColor       =   &H000000FF&
      Height          =   255
      Index           =   102
      Left            =   10680
      Top             =   8040
      Width           =   255
   End
   Begin VB.Shape Square 
      BorderColor     =   &H80000005&
      FillColor       =   &H000000FF&
      Height          =   255
      Index           =   101
      Left            =   11160
      Top             =   7680
      Width           =   255
   End
   Begin VB.Shape Square 
      BorderColor     =   &H80000005&
      FillColor       =   &H000000FF&
      Height          =   255
      Index           =   100
      Left            =   10680
      Top             =   7680
      Width           =   255
   End
   Begin VB.Shape Square 
      BorderColor     =   &H80000005&
      FillColor       =   &H000000FF&
      Height          =   255
      Index           =   93
      Left            =   11160
      Top             =   7320
      Width           =   255
   End
   Begin VB.Shape Square 
      BorderColor     =   &H80000005&
      FillColor       =   &H000000FF&
      Height          =   255
      Index           =   92
      Left            =   10680
      Top             =   7320
      Width           =   255
   End
   Begin VB.Shape Square 
      BorderColor     =   &H80000005&
      FillColor       =   &H000000FF&
      Height          =   255
      Index           =   91
      Left            =   11160
      Top             =   6960
      Width           =   255
   End
   Begin VB.Shape Square 
      BorderColor     =   &H80000005&
      FillColor       =   &H000000FF&
      Height          =   255
      Index           =   90
      Left            =   10680
      Top             =   6960
      Width           =   255
   End
   Begin VB.Shape Square 
      BorderColor     =   &H80000005&
      FillColor       =   &H000000FF&
      Height          =   255
      Index           =   83
      Left            =   11160
      Top             =   6600
      Width           =   255
   End
   Begin VB.Shape Square 
      BorderColor     =   &H80000005&
      FillColor       =   &H000000FF&
      Height          =   255
      Index           =   82
      Left            =   10680
      Top             =   6600
      Width           =   255
   End
   Begin VB.Shape Square 
      BorderColor     =   &H80000005&
      FillColor       =   &H000000FF&
      Height          =   255
      Index           =   81
      Left            =   11160
      Top             =   6240
      Width           =   255
   End
   Begin VB.Shape Square 
      BorderColor     =   &H80000005&
      FillColor       =   &H000000FF&
      Height          =   255
      Index           =   80
      Left            =   10680
      Top             =   6240
      Width           =   255
   End
   Begin VB.Shape Square 
      BorderColor     =   &H80000005&
      FillColor       =   &H000000FF&
      Height          =   255
      Index           =   73
      Left            =   11160
      Top             =   5880
      Width           =   255
   End
   Begin VB.Shape Square 
      BorderColor     =   &H80000005&
      FillColor       =   &H000000FF&
      Height          =   255
      Index           =   72
      Left            =   10680
      Top             =   5880
      Width           =   255
   End
   Begin VB.Shape Square 
      BorderColor     =   &H80000005&
      FillColor       =   &H000000FF&
      Height          =   255
      Index           =   71
      Left            =   11160
      Top             =   5520
      Width           =   255
   End
   Begin VB.Shape Square 
      BorderColor     =   &H80000005&
      FillColor       =   &H000000FF&
      Height          =   255
      Index           =   70
      Left            =   10680
      Top             =   5520
      Width           =   255
   End
   Begin VB.Shape Square 
      BorderColor     =   &H80000005&
      FillColor       =   &H000000FF&
      Height          =   255
      Index           =   63
      Left            =   11160
      Top             =   5160
      Width           =   255
   End
   Begin VB.Shape Square 
      BorderColor     =   &H80000005&
      FillColor       =   &H000000FF&
      Height          =   255
      Index           =   62
      Left            =   10680
      Top             =   5160
      Width           =   255
   End
   Begin VB.Shape Square 
      BorderColor     =   &H80000005&
      FillColor       =   &H000000FF&
      Height          =   255
      Index           =   61
      Left            =   11160
      Top             =   4800
      Width           =   255
   End
   Begin VB.Shape Square 
      BorderColor     =   &H80000005&
      FillColor       =   &H000000FF&
      Height          =   255
      Index           =   60
      Left            =   10680
      Top             =   4800
      Width           =   255
   End
   Begin VB.Shape Square 
      BorderColor     =   &H80000005&
      FillColor       =   &H000000FF&
      Height          =   255
      Index           =   53
      Left            =   11160
      Top             =   4440
      Width           =   255
   End
   Begin VB.Shape Square 
      BorderColor     =   &H80000005&
      FillColor       =   &H000000FF&
      Height          =   255
      Index           =   52
      Left            =   10680
      Top             =   4440
      Width           =   255
   End
   Begin VB.Shape Square 
      BorderColor     =   &H80000005&
      FillColor       =   &H000000FF&
      Height          =   255
      Index           =   51
      Left            =   11160
      Top             =   4080
      Width           =   255
   End
   Begin VB.Shape Square 
      BorderColor     =   &H80000005&
      FillColor       =   &H000000FF&
      Height          =   255
      Index           =   50
      Left            =   10680
      Top             =   4080
      Width           =   255
   End
   Begin VB.Shape Square 
      BorderColor     =   &H80000005&
      FillColor       =   &H000000FF&
      Height          =   255
      Index           =   43
      Left            =   11160
      Top             =   3720
      Width           =   255
   End
   Begin VB.Shape Square 
      BorderColor     =   &H80000005&
      FillColor       =   &H000000FF&
      Height          =   255
      Index           =   42
      Left            =   10680
      Top             =   3720
      Width           =   255
   End
   Begin VB.Shape Square 
      BorderColor     =   &H80000005&
      FillColor       =   &H000000FF&
      Height          =   255
      Index           =   41
      Left            =   11160
      Top             =   3360
      Width           =   255
   End
   Begin VB.Shape Square 
      BorderColor     =   &H80000005&
      FillColor       =   &H000000FF&
      Height          =   255
      Index           =   40
      Left            =   10680
      Top             =   3360
      Width           =   255
   End
   Begin VB.Shape Square 
      BorderColor     =   &H80000005&
      FillColor       =   &H000000FF&
      Height          =   255
      Index           =   33
      Left            =   11160
      Top             =   3000
      Width           =   255
   End
   Begin VB.Shape Square 
      BorderColor     =   &H80000005&
      FillColor       =   &H000000FF&
      Height          =   255
      Index           =   32
      Left            =   10680
      Top             =   3000
      Width           =   255
   End
   Begin VB.Shape Square 
      BorderColor     =   &H80000005&
      FillColor       =   &H000000FF&
      Height          =   255
      Index           =   31
      Left            =   11160
      Top             =   2640
      Width           =   255
   End
   Begin VB.Shape Square 
      BorderColor     =   &H80000005&
      FillColor       =   &H000000FF&
      Height          =   255
      Index           =   30
      Left            =   10680
      Top             =   2640
      Width           =   255
   End
   Begin VB.Shape Square 
      BorderColor     =   &H80000005&
      FillColor       =   &H000000FF&
      Height          =   255
      Index           =   23
      Left            =   11160
      Top             =   2280
      Width           =   255
   End
   Begin VB.Shape Square 
      BorderColor     =   &H80000005&
      FillColor       =   &H000000FF&
      Height          =   255
      Index           =   22
      Left            =   10680
      Top             =   2280
      Width           =   255
   End
   Begin VB.Shape Square 
      BorderColor     =   &H80000005&
      FillColor       =   &H000000FF&
      Height          =   255
      Index           =   21
      Left            =   11160
      Top             =   1920
      Width           =   255
   End
   Begin VB.Shape Square 
      BorderColor     =   &H80000005&
      FillColor       =   &H000000FF&
      Height          =   255
      Index           =   20
      Left            =   10680
      Top             =   1920
      Width           =   255
   End
   Begin VB.Shape Square 
      BorderColor     =   &H80000005&
      FillColor       =   &H000000FF&
      Height          =   255
      Index           =   13
      Left            =   11160
      Top             =   1560
      Width           =   255
   End
   Begin VB.Shape Square 
      BorderColor     =   &H80000005&
      FillColor       =   &H000000FF&
      Height          =   255
      Index           =   12
      Left            =   10680
      Top             =   1560
      Width           =   255
   End
   Begin VB.Shape Square 
      BorderColor     =   &H80000005&
      FillColor       =   &H000000FF&
      Height          =   255
      Index           =   11
      Left            =   11160
      Top             =   1200
      Width           =   255
   End
   Begin VB.Shape Square 
      BorderColor     =   &H80000005&
      FillColor       =   &H000000FF&
      Height          =   255
      Index           =   10
      Left            =   10680
      Top             =   1200
      Width           =   255
   End
   Begin VB.Shape Square 
      BorderColor     =   &H80000005&
      FillColor       =   &H000000FF&
      Height          =   255
      Index           =   3
      Left            =   11160
      Top             =   840
      Width           =   255
   End
   Begin VB.Shape Square 
      BorderColor     =   &H80000005&
      FillColor       =   &H000000FF&
      Height          =   255
      Index           =   2
      Left            =   10680
      Top             =   840
      Width           =   255
   End
   Begin VB.Shape Square 
      BorderColor     =   &H80000005&
      FillColor       =   &H000000FF&
      Height          =   255
      Index           =   1
      Left            =   11160
      Top             =   480
      Width           =   255
   End
   Begin VB.Shape Circ 
      BorderColor     =   &H80000005&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   0
      Left            =   6840
      Shape           =   3  'Circle
      Top             =   480
      Width           =   735
   End
   Begin VB.Shape Square 
      BorderColor     =   &H80000005&
      FillColor       =   &H000000FF&
      Height          =   255
      Index           =   0
      Left            =   10680
      Top             =   480
      Width           =   255
   End
   Begin VB.Shape Chooseback 
      BackColor       =   &H80000006&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000005&
      BorderWidth     =   3
      DrawMode        =   5  'Not Copy Pen
      Height          =   735
      Left            =   6600
      Top             =   9360
      Width           =   3735
   End
   Begin VB.Shape pinback 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      DrawMode        =   5  'Not Copy Pen
      FillStyle       =   0  'Solid
      Height          =   9015
      Left            =   10440
      Top             =   240
      Width           =   1215
   End
   Begin VB.Shape gameback 
      BackColor       =   &H80000006&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000005&
      BorderWidth     =   3
      DrawMode        =   5  'Not Copy Pen
      Height          =   9015
      Left            =   6600
      Top             =   240
      Width           =   3735
   End
   Begin VB.Shape Scoreback 
      BackColor       =   &H80000006&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000005&
      BorderWidth     =   3
      DrawMode        =   5  'Not Copy Pen
      Height          =   4695
      Left            =   240
      Top             =   5400
      Width           =   6135
   End
   Begin VB.Menu file 
      Caption         =   "File"
      Begin VB.Menu Newmen 
         Caption         =   "New Game"
         Begin VB.Menu SE 
            Caption         =   "Single Player - Easy"
         End
         Begin VB.Menu SN 
            Caption         =   "Single Player - Normal"
         End
         Begin VB.Menu SH 
            Caption         =   "Single Player - Hard"
         End
         Begin VB.Menu line 
            Caption         =   "--------------------------"
            Enabled         =   0   'False
         End
         Begin VB.Menu me 
            Caption         =   "2 Player - Easy"
         End
         Begin VB.Menu mn 
            Caption         =   "2 Player - Normal"
         End
         Begin VB.Menu mh 
            Caption         =   "2 Player - Hard"
         End
      End
      Begin VB.Menu Exitmen 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu hp 
      Caption         =   "Help"
      Begin VB.Menu Helpmen 
         Caption         =   "Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu Aboutmen 
         Caption         =   "About..."
      End
   End
End
Attribute VB_Name = "Master"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim counter(10), rannum(10) As Integer
Dim code(10), convert(10), test(10) As String
Dim diffcount, y, z, bestturn, turn, gametype, winloose As Integer
Dim diff, player1, player2, Bestname, guessname, codename, endcode, player As String
Dim returnval As Long
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long






'===================================== BUTTONS! ===============================================
'===================================== BUTTONS! ===============================================
Private Sub helpbut_Click()
Help.Show 'Shows help form
returnval = PlaySound("Windows XP Balloon", 0, 1)
End Sub
Private Sub exitbut_Click()
MsgBox ("Thanks for playing!")
End 'terminate program
End Sub

'Difficulty buttons
Private Sub S1_Click()
diff = 0
Call Difficulty
Call scorereset
Call names
Call Newgame

labturn2.Visible = False
labwin2.Visible = False
labname2.Visible = False
returnval = PlaySound("Windows XP Notify", 0, 1)
End Sub
Private Sub S2_Click()
diff = 1
Call Difficulty
Call scorereset
Call names
Call Newgame

labturn2.Visible = False
labwin2.Visible = False
labname2.Visible = False
returnval = PlaySound("Windows XP Notify", 0, 1)
End Sub
Private Sub S3_Click()
diff = 2
Call Difficulty
Call scorereset
Call names
Call Newgame

labturn2.Visible = False
labwin2.Visible = False
labname2.Visible = False
returnval = PlaySound("Windows XP Notify", 0, 1)
End Sub
Private Sub m1_Click()
diff = 3
Call Difficulty
Call scorereset
Call names
Call Newgame

labturn2.Visible = True
labwin2.Visible = True
labname2.Visible = True
returnval = PlaySound("Windows XP Notify", 0, 1)
End Sub
Private Sub m2_Click()
diff = 4
Call Difficulty
Call scorereset
Call names
Call Newgame

labturn2.Visible = True
labwin2.Visible = True
labname2.Visible = True
returnval = PlaySound("Windows XP Notify", 0, 1)
End Sub
Private Sub m3_Click()
diff = 5
Call Difficulty
Call scorereset
Call names
Call Newgame

labturn2.Visible = True
labwin2.Visible = True
labname2.Visible = True
returnval = PlaySound("Windows XP Notify", 0, 1)
End Sub
'/Change Difficulty

'MENU BUTTONS
Private Sub Exitmen_Click()
Call exitbut_Click
End Sub
Private Sub Helpmen_Click()
Call helpbut_Click
End Sub
Private Sub SE_Click()
Call S1_Click
End Sub
Private Sub Sn_Click()
Call S2_Click
End Sub
Private Sub Sh_Click()
Call S3_Click
End Sub
Private Sub me_Click()
Call m1_Click
End Sub
Private Sub mn_Click()
Call m2_Click
End Sub
Private Sub mh_Click()
Call m3_Click
End Sub
Private Sub Aboutmen_Click()
About.Show
End Sub
'/MENU BUTTONS!

'===================================== /BUTTONS!!!!!! ========================================================
'===================================== /BUTTONS!!!!!! =========================================================














'===================================== SCORE BOARD FUNCTIONS =====================================================
'===================================== SCORE BOARD FUNCTIONS =====================================================
Private Sub scoreboard() 'Adds the names to the scoreboard

labname1 = player1
labname2 = player2

If gametype = 1 Then
labguess = guessname & "'s Turn To Guess"
Else
labguess = "Mastermind Scoreboard"
End If

End Sub

Private Sub names()

'1 player
If player1 = "" Then
    player1 = InputBox("Player 1 Please Enter Your Name") 'player enters name
    player1 = UCase(Left(player1, 1)) & Mid(player1, 2) 'Converts the first letter of input to capitals, for neatness.
End If
If player1 = "" Then 'if the player didnt input a name this checks for that which is then fixed in namecheck.
    player = "n/a"
End If
    
'2 player
If gametype = 1 Then
    If player2 = "" Then
        player2 = InputBox("Player 2 Please Enter Your Name")
        player2 = UCase(Left(player2, 1)) & Mid(player2, 2)
    End If
    If player2 = "" Then
        player = "n/a"
    End If
End If

If gametype = 0 Then
    Bestname = player1 'Used for the best game for single player, since it doesnt have another person to compete with.
End If

guessname = player1 'same as bestname ^
codename = player2 'same as bestname ^
Call namecheck
Call scoreboard
End Sub

Private Sub namecheck()

'Checks for missing player name, then disallows player if its not found and tells the player to input.
If player = "n/a" Then
    If gametype = 0 Then
        MsgBox ("Please enter name.")
    Else
        MsgBox ("Please enter both names.")
    End If
Master.Width = 6585 'Menu size
Master.Height = 6015
player = ""
Else
Master.Height = 11120 'Full game screen
Master.Width = 12000
End If

End Sub
'=================================== /SCORE BOARD FUNCTIONS ======================================================-
'=================================== /SCORE BOARD FUNCTIONS ======================================================













'================================= NEW GAME FUNCTIONS ===========================================================
'================================= NEW GAME FUNCTIONS ===========================================================
Private Sub scorereset() 'self explanitory
labwin1 = 0
labwin2 = 0
labturn1 = 0
labturn2 = 0
gamenum = 0
bestturn = 0
Bestname = ""
msg = ""
bestnum = ""
End Sub

Private Sub Newgame()
'Starts a new game
y = 0 'Resets y and counter

'blanks all the circles
For z = 0 To 11
    For x = 0 To 3
        Circ(z & x).FillColor = &H80000007 'Black
    Next
Next

'blanks all the squares
For z = 0 To 11
    For x = 0 To 3
        Square(z & x).FillStyle = 1
    Next
Next

'Greens all the Pics
For x = 0 To 3
    Pic(x).BackColor = &HFF00& 'Green
    counter(x) = 0
Next

'Resets all the Pic squares to greens position
For Index = 0 To diffcount
counter(Index) = 1
Next

Pick.Caption = "Pick Colours!" 'changes button back to pick colours and disables new game in menu

If gametype = 1 Then
Pick.Caption = "Accept Code"
msg = guessname & ", please look away while " & codename & " chooses the code."
Else
Pick.Caption = "New Game!"
End If

End Sub
'============================================ /NEW GAME FUNCTIONS ===========================================
'============================================ /NEW GAME FUNCTIONS ===========================================












'============================================ GAME FUNCTIONS =================================================
'============================================ GAME FUNCTIONS =================================================
Private Sub Difficulty()

'Determines the difficulty of the game.
Select Case diff
    Case "0"
        Master.Caption = "Mastermind - Single Player - Easy"
        gametype = 0
        
    Case "1"
        Master.Caption = "Mastermind - Single Player - Normal"
        gametype = 0
        
    Case "2"
        Master.Caption = "Mastermind - Single Player - Hard"
        gametype = 0
        
    Case "3"
        Master.Caption = "Mastermind - 2 Player - Easy"
        diff = 0 'sets game to easy
        gametype = 1
        
    Case "4"
        Master.Caption = "Mastermind - 2 Player - Normal"
        diff = 1 'sets game to normal
        gametype = 1
        
    Case "5"
        Master.Caption = "Mastermind - 2 Player - Hard"
        diff = 2 'sets game to Hard
        gametype = 1
End Select

End Sub

Private Sub Pic_Mouseup(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

'Selects difficulty
Select Case diff
Case 0
diffcount = 4
Case 1
diffcount = 6
Case 2
diffcount = 8
End Select

'Changes the colours
If counter(Index) <= diffcount Then

    'counter for left click
    If Button = 1 Then
        counter(Index) = counter(Index) + 1
    End If
    
    'counter for right click
    If Button = 2 Then
        counter(Index) = counter(Index) - 1
    End If
    
    'Checks for reset in left click
    If Button = 1 Then
        If counter(Index) > diffcount Then
            counter(Index) = 1
        End If
    End If
    
    'Checks for reset in a right click
    If Button = 2 Then
        If counter(Index) <= 0 Then
            counter(Index) = diffcount
        End If
    End If
    
    Select Case counter(Index)
        'Easy
        Case Is = 1
            Pic(Index).BackColor = &HFF00& 'Green
                
        Case Is = 2
            Pic(Index).BackColor = &HFF& 'Red
                
        Case Is = 3
            Pic(Index).BackColor = &H80FF& 'Orange
                
        Case Is = 4
            Pic(Index).BackColor = &HFFFF& 'Yellow
                
        'Normal
        Case Is = 5
            Pic(Index).BackColor = &HFFFF00 'cyan
                
        Case Is = 6
            Pic(Index).BackColor = &HFF00FF 'pink
                
        'Hard
        Case Is = 7
            Pic(Index).BackColor = &HFF0000 'blue
        Case Is = 8
            Pic(Index).BackColor = &HFFFFFF 'white
    End Select
End If

End Sub


'PICK BUTTON
Private Sub Pick_Click()

Call conv2

'Turns
If Pick.Caption = "Pick Colours!" Then

y = y + 1 'Each click adds to the next row
 
    'places the picked colours onto the game screen, X is across going (0 to 3) in the array and Y is Down (0 to 11)
    For x = 0 To 3 'circ is circles, pic is the 4 chosen colours. -1 is to counter the starting at 1 problem, y & x allows me to make it seem like there is 2 arrays
        Circ(-1 + y & x).FillColor = Pic(x).BackColor
    Next
                
               
    Call codes
    Call squares
    
            
    'Add to Turns
    If turn = 0 Then
        labturn1 = labturn1 + 1
    End If
            
    If turn = 1 Then
        labturn2 = labturn2 + 1
    End If
    ' /Add to Turns
        
    Call conv2
    Call win
End If
'/Turns


'NEW GAME FOR 2 PLAYER
If gametype = 1 Then
    If Pick.Caption = "Accept Code" Then
        msg = ""
        If y = 0 Then
            Call codes
        End If
        
        'Greens all the Pics
        For x = 0 To 3
            Pic(x).BackColor = &HFF00& 'Green
            counter(x) = 1
        Next
        
        Pick.Caption = "Pick Colours!"
        returnval = PlaySound("notify", 0, &H0)
    End If
End If
'/NEW GAME FOR 2 PLAYER


'New game button
'If the button is new game! then call the newgame button, if not it does the loop to place colours in circles
If Pick.Caption = "New Game!" Then
    Call Newgame
    If gametype = 0 Then
        Call generatecode
        Call conv1
        Call codes
        Pick.Caption = "Pick Colours!"
    End If
End If
'/New game button


'No turns left
If y = 12 Then
Call gameover
winloose = 0
End If
'/No turns left

'/PICK BUTTON
End Sub

'GAME WIN
Private Sub win()

u = 0
For x = 0 To 3

    'Check for win
    If convert(x) = test(x) Then 'Checks for a win, if u gets to 4 its a win
        u = u + 1
    End If
    '/Check for win
    
    If u = 4 Then
        
        'Determines the best game
        If bestturn = 0 Then
            Bestname = guessname
            bestturn = y
            If bestturn = 1 Then
                bestnum = Bestname & " - " & bestturn & " Turn."
            Else
                bestnum = Bestname & " - " & bestturn & " Turns."
            End If
        End If
        
        If y < bestturn Then
            Bestname = guessname
            bestturn = y
            If bestturn = 1 Then
                bestnum = Bestname & " - " & bestturn & " Turn."
            Else
                bestnum = Bestname & " - " & bestturn & " Turns."
            End If
        End If
        
        'Single Player
        If gametype = 0 Then
        bestnum = player1 & " - " & bestturn & " Turns"
        End If
        '/Single Player
        
        
        '/Determines the best game
        
        
        If turn = 1 Then
            labwin2 = labwin2 + 1 'Scoreboard
        Else
            labwin1 = labwin1 + 1 'Scoreboard
        End If
        y = 12
        winloose = 1
        Call gameover
    End If
Next

End Sub

'GAME OVER
Private Sub gameover()

If y = 12 Then 'If the y value is greater then 12 change button to new game and tell user its over.
    
    gamenum = gamenum + 1
    Pick.Caption = "New Game!"
    
    
    'Displays Code at end of game
    For x = 0 To 4
        endcode = endcode & test(x) 'Adds the 4 letters of code to a single string
    Next
    
    If winloose = 0 Then
        MsgBox ("Sorry, the Code Was " & endcode & "!!!")
    Else
        MsgBox ("Good Job, You Win!!!")
    End If
    
    winloose = 0
    endcode = "" 'Clears code string for next game
    '/Displays Code at end of game
    
    
    'Change Players Turn
    If gametype = 1 Then
        If turn = 0 Then 'Flips turns for 2 player
            turn = 1
            guessname = player2
            codename = player1
        Else
            turn = 0 'TURN 1 IS PLAYER 2
            guessname = player1
            codename = player2
        End If
        Call scoreboard
    End If
    '/Change Players Turn
    
End If
y = 0

'/GAME OVER
End Sub

Private Sub conv1()

'CONVERT NUMBERS INTO CODE FOR SINGLE PLAYER

'This code converts the random numbers generated in generate code to letters, which can then be used for the code.
For x = 0 To 3
    If rannum(x) = 1 Then
        convert(x) = "R"
    End If
    If rannum(x) = 2 Then
        convert(x) = "O"
    End If
    If rannum(x) = 3 Then
        convert(x) = "Y"
    End If
    If rannum(x) = 4 Then
        convert(x) = "G"
    End If
    If rannum(x) = 5 Then
        convert(x) = "C"
    End If
    If rannum(x) = 6 Then
        convert(x) = "P"
    End If
    If rannum(x) = 7 Then
        convert(x) = "B"
    End If
    If rannum(x) = 8 Then
        convert(x) = "W"
    End If
Next
'/CONVERT NUMBERS INTO CODE FOR SINGLE PLAYER

End Sub

Private Sub conv2()
'COVERTS COLOURS TO CODE FOR 2 PLAYER

'This code translates the colour of the 4 colour boxes into letter form, which can be translated to code
For x = 0 To 3
    If Pic(x).BackColor = &HFF& Then 'Red
        convert(x) = "R"
    End If
    If Pic(x).BackColor = &H80FF& Then 'orange
        convert(x) = "O"
    End If
    If Pic(x).BackColor = &HFFFF& Then 'yellow
        convert(x) = "Y"
    End If
    If Pic(x).BackColor = &HFF00& Then 'green
        convert(x) = "G"
    End If
    If Pic(x).BackColor = &HFFFF00 Then 'cyan
        convert(x) = "C"
    End If
    If Pic(x).BackColor = &HFF00FF Then 'pink
        convert(x) = "P"
    End If
    If Pic(x).BackColor = &HFF0000 Then 'blue
        convert(x) = "B"
    End If
    If Pic(x).BackColor = &HFFFFFF Then 'white
        convert(x) = "W"
    End If
Next

'/COVERTS COLOURS TO CODE FOR 2 PLAYER
End Sub

'USED FOR RESETTING THE CODE FOR EACH TIME SQUARES IS CALLED
Private Sub codes()

If y = 0 Then
    For x = 0 To 3
        test(x) = convert(x) 'Test is a permanent record of the code, convert is the original extract
    Next
    For x = 0 To 3
        code(x) = test(x) 'Code is the temporary code used to determine right or wrong colour combinations
        
    Next
End If

If y > 0 Then 'This code resets the temporary code to the original after each attempt at breaking the code.
    For x = 0 To 3
        code(x) = test(x)
    Next
End If

'/USED FOR RESETTING THE CODE FOR EACH TIME SQUARES IS CALLED
End Sub

Private Sub squares()

'COMAPRES GUESS TO CODE AND PLACES THE ANSWER TO THE SQUARES
    z = 0
    For x = 0 To 3
        If convert(x) = code(x) Then 'This loop finds the right colour right place circles and removes them from the temp code
            Square(-1 + y & z).FillStyle = 0 ' "-1 +y & z" is the position of the squares on the right
            z = z + 1
            code(x) = 0 'This removes the posibility of the same code being found again by the next loop
            convert(x) = 1
        End If
    Next
        
    For x = 0 To 3 'This loop follows the first one and finds all the right colour wrong place circles
        If convert(0) = code(x) Then 'looks through each individual array and finds the colours in the entire guess array
            Square(-1 + y & z).FillStyle = 7
            z = z + 1
            code(x) = 0
            convert(0) = 1
        End If
        If convert(1) = code(x) Then
            Square(-1 + y & z).FillStyle = 7
            z = z + 1
            code(x) = 0
            convert(1) = 1
        End If
        If convert(2) = code(x) Then
            Square(-1 + y & z).FillStyle = 7
            z = z + 1
            code(x) = 0
            convert(2) = 1
        End If
        If convert(3) = code(x) Then
            Square(-1 + y & z).FillStyle = 7
            z = z + 1
            code(x) = 0
            convert(3) = 1
        End If
    Next
    
'/COMAPRES GUESS TO CODE AND PLACES THE ANSWER TO THE SQUARES
End Sub

Private Sub generatecode()
'GENERATE RANDOM NUMBERS FOR SINGLE PLAYER

'EASY
If diff = 0 Then
    Randomize 'Makes sure each rnd is different
    For x = 0 To 3
        rannum(x) = Int(4 * Rnd(x) + 1) 'generates 4 random numbers between 1 and 4 for easy
    Next
End If

'NORMAL
If diff = 1 Then
    Randomize
    For x = 0 To 3
        rannum(x) = Int(6 * Rnd(x) + 1) 'generates 4 random numbers between 1 and 6 for normal
    Next
End If

'HARD
If diff = 2 Then
    Randomize
    For x = 0 To 3
        rannum(x) = Int(8 * Rnd(x) + 1) 'generates 4 random numbers between 1 and 8 for hard
    Next
End If

'/GENERATE RANDOM NUMBERS FOR SINGLE PLAYER
End Sub
'============================================ GAME FUNCTIONS =================================================
'============================================ GAME FUNCTIONS =================================================
