VERSION 5.00
Begin VB.Form frmHelp 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00800000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "How to Play"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8550
   Icon            =   "Help.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   8550
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00404040&
      Caption         =   "RETURN"
      BeginProperty Font 
         Name            =   "IndustryIncW00-Base"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6600
      Width           =   2535
   End
   Begin VB.Label lblLength 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "[TWO UNITS]"
      BeginProperty Font 
         Name            =   "Purista"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   6840
      TabIndex        =   16
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label lblLength 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "[THREE UNITS]"
      BeginProperty Font 
         Name            =   "Purista"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   5160
      TabIndex        =   15
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label lblLength 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "[THREE UNITS]"
      BeginProperty Font 
         Name            =   "Purista"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   3480
      TabIndex        =   14
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label lblLength 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "[FOUR UNITS]"
      BeginProperty Font 
         Name            =   "Purista"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   13
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label lblLength 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "[FIVE UNITS]"
      BeginProperty Font 
         Name            =   "Purista"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label lblPShip 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "AIRCRAFT CARRIER"
      BeginProperty Font 
         Name            =   "Purista"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label lblPShip 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "BATTLECRUISER"
      BeginProperty Font 
         Name            =   "Purista"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   10
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label lblPShip 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SUBMARINE"
      BeginProperty Font 
         Name            =   "Purista"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   3480
      TabIndex        =   9
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label lblPShip 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CRUISER"
      BeginProperty Font 
         Name            =   "Purista"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   5160
      TabIndex        =   8
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label lblPShip 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "FRIGATE"
      BeginProperty Font 
         Name            =   "Purista"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   6840
      TabIndex        =   7
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label lblSubtitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "- STRATEGY REDEFINED -"
      BeginProperty Font 
         Name            =   "Purista"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Top             =   840
      Width           =   4455
   End
   Begin VB.Label lblLogo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "BATTLESHIP"
      BeginProperty Font 
         Name            =   "IndustryIncW00-Base"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   855
      Left            =   2040
      TabIndex        =   5
      Top             =   120
      Width           =   4455
   End
   Begin VB.Image imgPShip 
      Height          =   1095
      Index           =   0
      Left            =   120
      Picture         =   "Help.frx":5C12
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Image imgPShip 
      Height          =   1095
      Index           =   1
      Left            =   1800
      Picture         =   "Help.frx":15976
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Image imgPShip 
      Height          =   1095
      Index           =   2
      Left            =   3480
      Picture         =   "Help.frx":256DA
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Image imgPShip 
      Height          =   1095
      Index           =   3
      Left            =   5160
      Picture         =   "Help.frx":3543E
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Image imgPShip 
      Height          =   1095
      Index           =   4
      Left            =   6840
      Picture         =   "Help.frx":451A2
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label lblIntro 
      BackColor       =   &H00400000&
      BackStyle       =   0  'Transparent
      Caption         =   "INTRO PARAGRAPH "
      BeginProperty Font 
         Name            =   "Purista"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   8295
   End
   Begin VB.Label lblDesc1 
      BackColor       =   &H00400000&
      BackStyle       =   0  'Transparent
      Caption         =   "DESCRIPTION PARAGRAPH 1"
      BeginProperty Font 
         Name            =   "Purista"
         Size            =   9.75
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
      Top             =   1920
      Width           =   8295
   End
   Begin VB.Label lblDesc2 
      BackColor       =   &H00400000&
      BackStyle       =   0  'Transparent
      Caption         =   "DESCRIPTION PARAGRAPH 2"
      BeginProperty Font 
         Name            =   "Purista"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   4320
      Width           =   8295
   End
   Begin VB.Label lblDesc3 
      BackColor       =   &H00400000&
      BackStyle       =   0  'Transparent
      Caption         =   "DESCRIPTION PARAGRAPH 3"
      BeginProperty Font 
         Name            =   "Purista"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   5160
      Width           =   8295
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdReturn_Click()
    
    Unload Me
    
End Sub

Private Sub Form_Load()
        
    Dim Msg As String
    
    Me.BackColor = RGB(40, 60, 70)
    cmdReturn.BackColor = CMDBOXCOLOR
    
    Msg = "Battleship is a guessing board game for two players. "
    Msg = Msg & "The objective of Battleship is to be the first player to sink "
    Msg = Msg & "all 5 of your opponent's ships."
    lblIntro.Caption = Msg
    
    Msg = "When starting up a game, there will be two grids in which "
    Msg = Msg & "ships can be placed on and later be fired at. "
    Msg = Msg & "You are to place all of the following ships on your "
    Msg = Msg & "grid before you can begin the game:"
    lblDesc1.Caption = Msg
    
    Msg = "If you think that individually placing ships is too much of "
    Msg = Msg & "a hassle, then otherwise simply click 'Random Placement' "
    Msg = Msg & "to automatically place all 5 ships onto the board. If you do not "
    Msg = Msg & "prefer the current placements, then you can click on the 'Reset Positions' "
    Msg = Msg & "button to reset the grid."
    lblDesc2.Caption = Msg
    
    Msg = "Once you are satisfied with your placements, click the 'Begin Operation' button to "
    Msg = Msg & "begin the game. There will be a 50-50 chance that the computer player will begin "
    Msg = Msg & "shooting first. Click on the computer's grid to fire on the computer's ships. "
    Msg = Msg & "A red box means a miss, while a green box means that you have hit "
    Msg = Msg & "one of the computer's ships. If you have sunk all of the computer's ships, then that "
    Msg = Msg & "will constitute as a win for you."
    lblDesc3.Caption = Msg
    
    CentreForm Me
    
End Sub
