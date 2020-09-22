VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Custom Splash Screen Demo Project"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   8460
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboSpeed 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6600
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   4320
      Width           =   1215
   End
   Begin VB.ComboBox cboFade 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   4320
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Show the form without the close timer set"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   2
      ToolTipText     =   "Double click the splash form to unload"
      Top             =   5040
      Width           =   3855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&That was cool, show it to me again!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   5040
      Width           =   3855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fade Speed:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5160
      TabIndex        =   6
      Top             =   4440
      Width           =   1365
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fade Styles:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   600
      TabIndex        =   4
      Top             =   4440
      Width           =   1320
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmMain.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   8055
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private dsaFade As dlgShowActions

Private Sub cboFade_Click()
    dsaFade = cboFade.ItemData(cboFade.ListIndex)
End Sub

Private Sub Command1_Click()

    With frmSplash
        .Duration = 8
        .FadeSpeed = Val(cboSpeed.Text)
        .DialogAction dsaFade
        .Fixed = False
        .Show
    End With
    
End Sub

Private Sub Command2_Click()
    With frmSplash
        .Duration = 0 ' Don't use the timer to unload
        .FadeSpeed = Val(cboSpeed.Text)
        .DialogAction dsaFade
        .Fixed = False
        .Show
    End With
End Sub

Private Sub Form_Load()
    Me.Left = (Screen.Width / 2) - (Me.Width / 2)
    Me.Top = (Screen.Height / 2) - (Me.Height / 2)
    
    ' Load a combo with the different fade options
    With cboFade
        .Clear
        .AddItem "Show Normally": .ItemData(.NewIndex) = 0
        .AddItem "Fade In, Close Normally": .ItemData(.NewIndex) = 1
        .AddItem "Open Normally, Fade Out": .ItemData(.NewIndex) = 2
        .AddItem "Fade In and Fade Out": .ItemData(.NewIndex) = 3
        .Text = .List(3)
        dsaFade = .ItemData(3)
    End With
    
    ' Load a combo with different fade speeds
    Dim intSpeed As Integer
    With cboSpeed
        .Clear
        For intSpeed = 20 To 200 Step 10
            .AddItem intSpeed
        Next
        .Text = frmSplash.FadeSpeed
    End With
    
    
End Sub






