VERSION 5.00
Begin VB.Form frmExit 
   BackColor       =   &H00400000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exit Game"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6870
   Icon            =   "Exit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   6870
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNo 
      BackColor       =   &H00404040&
      Caption         =   "NO"
      BeginProperty Font 
         Name            =   "IndustryIncW00-Base"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      MaskColor       =   &H00FF8080&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2160
      Width           =   3255
   End
   Begin VB.CommandButton cmdYes 
      BackColor       =   &H00404040&
      Caption         =   "YES"
      BeginProperty Font 
         Name            =   "IndustryIncW00-Base"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      MaskColor       =   &H00FF8080&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2160
      Width           =   3255
   End
   Begin VB.Image imgFlag 
      Height          =   900
      Index           =   0
      Left            =   720
      Picture         =   "Exit.frx":5C12
      Top             =   960
      Width           =   1350
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H00400000&
      BackStyle       =   0  'Transparent
      Caption         =   "Are you sure you want to exit Battleship?"
      BeginProperty Font 
         Name            =   "Purista"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Index           =   0
      Left            =   2280
      TabIndex        =   3
      Top             =   960
      Width           =   4095
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      BackStyle       =   0  'Transparent
      Caption         =   "EXIT GAME?"
      BeginProperty Font 
         Name            =   "IndustryIncW00-Base"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6615
   End
End
Attribute VB_Name = "frmExit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdNo_Click()
    
    Unload Me
    
End Sub

Private Sub cmdYes_Click()
    
    End
    
End Sub

Private Sub Form_Load()
    
    Me.BackColor = RGB(40, 60, 70)
    cmdYes.BackColor = CMDBOXCOLOR
    cmdNo.BackColor = CMDBOXCOLOR
        
    CentreForm Me
    
End Sub

