VERSION 5.00
Begin VB.Form frmMenu 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Battleship 2016 - 2.0 Alpha Build"
   ClientHeight    =   8070
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   14055
   Icon            =   "Menu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "Menu.frx":5C12
   ScaleHeight     =   8070
   ScaleWidth      =   14055
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrTime 
      Interval        =   1000
      Left            =   13440
      Top             =   1200
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00404040&
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "IndustryIncW00-Base"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10560
      MaskColor       =   &H00FF8080&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7080
      Width           =   3375
   End
   Begin VB.CommandButton cmdOptions 
      BackColor       =   &H00404040&
      Caption         =   "OPTIONS"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "IndustryIncW00-Base"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7080
      Width           =   3375
   End
   Begin VB.CommandButton cmdStats 
      BackColor       =   &H00404040&
      Caption         =   "STATS"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "IndustryIncW00-Base"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7080
      Width           =   3375
   End
   Begin VB.CommandButton cmdPlay 
      BackColor       =   &H00404040&
      Caption         =   "PLAY"
      BeginProperty Font 
         Name            =   "IndustryIncW00-Base"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7080
      Width           =   3375
   End
   Begin VB.Label lblDialog 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "BATTLESHIP 2016 - WELCOME TO THE NEXT GENERATION OF NAVAL WARFARE"
      BeginProperty Font 
         Name            =   "Purista"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   7680
      Width           =   13815
   End
   Begin VB.Label lblTime 
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Purista"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   615
      Left            =   12240
      TabIndex        =   8
      Top             =   480
      Width           =   1695
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
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   4215
   End
   Begin VB.Label lblLogo2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SHIP"
      BeginProperty Font 
         Name            =   "IndustryIncW00-Base"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   855
      Left            =   2520
      TabIndex        =   6
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "VERSION 2.0 ALPHA "
      BeginProperty Font 
         Name            =   "Purista"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   11640
      TabIndex        =   5
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label lblLogo 
      BackStyle       =   0  'Transparent
      Caption         =   "BATTLE"
      BeginProperty Font 
         Name            =   "IndustryIncW00-Base"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
   Begin VB.Menu mnuGame 
      Caption         =   "Game"
      Begin VB.Menu mnuNewGame 
         Caption         =   "New Game"
      End
      Begin VB.Menu mnuLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuInstruct 
         Caption         =   "How to Play"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSplash 
         Caption         =   "Splash"
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Title: BATTLESHIP 2016
'Author: Yun Jie (Jeffrey) Li
'Date: May 19, 2016
'Files: Battleship.bas, About.frm, Exit.frm, Game.frm, GameOver.frm, Help.frm, Return.frm,
'       Splash.frm, Theatres.frm, respective .frx files.
'Purpose: The purpose of this application is to recreate the classical board game
'         'Battleship'. Also for bragging rights.

Const FORM_WIDTH = 14145
Const FORM_HEIGHT = 8790

Option Explicit

Private Sub cmdExit_Click()
    
    mnuExit_Click
    
End Sub

Private Sub cmdPlay_Click()
        
    TheatreFromMenu = True
    frmTheatres.Show vbModal
    
End Sub

Private Sub Form_Load()
    
    PassiveClose = False
    
    SetColours
    lblTime.Caption = Time$()
    
    With Me
        .Width = FORM_WIDTH
        .Height = FORM_HEIGHT
    End With
    CentreForm Me
    
End Sub

Private Sub SetColours()

    lblLogo.ForeColor = RGB(100, 140, 160)
    cmdPlay.BackColor = CMDBOXCOLOR
    cmdOptions.BackColor = CMDBOXCOLOR
    cmdStats.BackColor = CMDBOXCOLOR
    cmdExit.BackColor = CMDBOXCOLOR

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    lblDialog = "BATTLESHIP 2016 - WELCOME TO THE NEXT GENERATION OF NAVAL WARFARE"
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
'    If Not PassiveClose Then
'        If UnloadMode = 0 Then
'            Cancel = 1
'        End If
'        mnuExit_Click
'    End If
    
End Sub

Private Sub mnuAbout_Click()
    
    frmAbout.Show vbModal
    
End Sub

Private Sub mnuExit_Click()
    
    EndProgram
    
End Sub

Private Sub mnuInstruct_Click()
    
    frmHelp.Show vbModal
    
End Sub

Private Sub mnuNewGame_Click()
    
    cmdPlay_Click
    
End Sub

Private Sub mnuSplash_Click()
    
    PassiveClose = True
    frmSplash.Show
    Unload Me
    
End Sub

Private Sub tmrTime_Timer()
    
    lblTime.Caption = Time$()
    
End Sub
