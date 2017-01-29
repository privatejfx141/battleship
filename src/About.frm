VERSION 5.00
Begin VB.Form frmAbout 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Battleship 2016 - About"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8655
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "About.frx":5C12
   ScaleHeight     =   4710
   ScaleWidth      =   8655
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
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4200
      Width           =   2415
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "INSERT DESCRIPTION HERE"
      BeginProperty Font 
         Name            =   "Purista"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2535
      Left            =   3960
      TabIndex        =   3
      Top             =   1320
      Width           =   4335
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
      Left            =   3960
      TabIndex        =   1
      Top             =   120
      Width           =   4335
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
      Left            =   3960
      TabIndex        =   0
      Top             =   840
      Width           =   4335
   End
End
Attribute VB_Name = "frmAbout"
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

Option Explicit

Private Sub cmdReturn_Click()
    
    Unload Me
    
End Sub

Private Sub Form_Load()
    
    Dim Msg As String
    
    Me.BackColor = RGB(40, 60, 70)
    lblLogo.ForeColor = RGB(60, 90, 100)
    cmdReturn.BackColor = CMDBOXCOLOR
    
    Msg = "BATTLESHIP 2016 allows you to play the classic Battleship board game against the computer. "
    Msg = Msg & "For instructions on how to play the game, click on the 'How to Play' button under 'Help'."
    Msg = Msg & vbCrLf & vbCrLf & "Programmed by Jeffrey Li 12H,"
    Msg = Msg & vbCrLf & "Riverdale Collegiate Institute."
    
    lblInfo.Caption = Msg
    
    CentreForm Me
    
End Sub

