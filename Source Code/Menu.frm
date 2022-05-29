VERSION 5.00
Begin VB.Form Menu 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Main Menu"
   ClientHeight    =   3135
   ClientLeft      =   6045
   ClientTop       =   3900
   ClientWidth     =   4830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Menu.frx":0000
   ScaleHeight     =   3135
   ScaleWidth      =   4830
   Begin VB.TextBox player2 
      Height          =   285
      Left            =   3960
      TabIndex        =   13
      Top             =   2760
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Player1 
      Height          =   285
      Left            =   3240
      TabIndex        =   12
      Top             =   2760
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Diff 
      Height          =   285
      Left            =   3960
      TabIndex        =   11
      Top             =   2400
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton m3 
      Caption         =   "Hard"
      Height          =   375
      Left            =   3120
      TabIndex        =   8
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton m2 
      Caption         =   "Normal "
      Height          =   375
      Left            =   3120
      TabIndex        =   7
      Top             =   1680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton m1 
      Caption         =   "Easy"
      Height          =   375
      Left            =   3120
      TabIndex        =   6
      Top             =   1200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton S3 
      Caption         =   "Hard"
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton S2 
      Caption         =   "Normal "
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   1680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton S1 
      BackColor       =   &H00000000&
      Caption         =   "Easy"
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   1200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton quit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   1800
      TabIndex        =   10
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton How 
      Caption         =   "How to Play"
      Height          =   375
      Left            =   1800
      TabIndex        =   9
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton m0 
      Caption         =   "2 Player"
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton S0 
      Caption         =   "Single Player"
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Mastermind"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   42
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4815
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Quit_Click() 'Exit Button
End
End Sub

Private Sub how_Click() 'Brings up Help screen
Help.Show 'Help form
End Sub

Private Sub S0_Click() 'shows the difficulty buttons for single player
S1.Visible = True
S2.Visible = True
S3.Visible = True
m1.Visible = False '1=easy, 2=normal, 3=hard, s=single player, m=2 player
m2.Visible = False
m3.Visible = False
End Sub

Private Sub m0_Click() 'shows the difficulty buttons for 2 player
m1.Visible = True
m2.Visible = True
m3.Visible = True
S1.Visible = False
S2.Visible = False
S3.Visible = False
End Sub


'========================Difficulties=============================
Private Sub S1_Click()
Diff.Text = 0
Main.Show
Menu.Hide 'shows the game screen, hides the menu
End Sub
Private Sub S2_Click()
Diff.Text = 1
Main.Show
Menu.Hide
End Sub
Private Sub S3_Click()
Diff.Text = 2
Main.Show
Menu.Hide
End Sub
Private Sub m1_Click()
Diff.Text = 3
Player1 = InputBox("Player 1 Enter Name") 'asks for player1's name for use when taking turns
player2 = InputBox("Player 2 Enter Name") ''
If Player1 <> "" Then
    If player2 <> "" Then
        Main.Show
        Menu.Hide
    Else
        MsgBox ("Please Enter your Names")
    End If
Else
    MsgBox ("Please Enter your Names")
End If
End Sub
Private Sub m2_Click()
Diff.Text = 4
Player1 = InputBox("Player 1 Enter Name")
player2 = InputBox("Player 2 Enter Name")
If Player1 <> "" Then
    If player2 <> "" Then
        Main.Show
        Menu.Hide
    Else
        MsgBox ("Please Enter your Names")
    End If
Else
    MsgBox ("Please Enter your Names")
End If
End Sub
Private Sub m3_Click()
Diff.Text = 5
Player1 = InputBox("Player 1 Enter Name")
player2 = InputBox("Player 2 Enter Name")
If Player1 <> "" Then
    If player2 <> "" Then
        Main.Show
        Menu.Hide
    Else
        MsgBox ("Please Enter your Names")
    End If
Else
    MsgBox ("Please Enter your Names")
End If
End Sub
'========================Difficulties=============================
