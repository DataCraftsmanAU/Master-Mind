VERSION 5.00
Begin VB.Form About 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About Mastermind"
   ClientHeight    =   2925
   ClientLeft      =   5580
   ClientTop       =   3735
   ClientWidth     =   4530
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "About.frx":0E42
   ScaleHeight     =   2925
   ScaleWidth      =   4530
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Supervised by: Darrin Inwood"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   0
      TabIndex        =   3
      Top             =   2160
      Width           =   4455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Caitlin Purcell"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   2040
      TabIndex        =   2
      Top             =   1560
      Width           =   4215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Created by: Michael Jennings"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Mastermind"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   855
      Left            =   720
      MousePointer    =   2  'Cross
      TabIndex        =   0
      Top             =   0
      Width           =   3255
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim counter As Integer

Private Sub Label1_Click()
About.Picture = Fake.Picture ' I AM AN EASTER EGG!
If counter = 0 Then
Label2.Visible = False
Label3.Visible = False
Label4.Visible = False
counter = 1
Else
Label2.Visible = True
Label3.Visible = True
Label4.Visible = True
counter = 0
End If
End Sub

Private Sub Label2_Click()

End Sub
