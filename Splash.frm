VERSION 5.00
Begin VB.Form frmSplash 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Battleship"
   ClientHeight    =   5175
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8895
   Icon            =   "Splash.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Splash.frx":030A
   ScaleHeight     =   5175
   ScaleWidth      =   8895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrTimer 
      Interval        =   2000
      Left            =   8400
      Top             =   4320
   End
   Begin VB.Label lblLoading 
      BackStyle       =   0  'Transparent
      Caption         =   "Loading..."
      BeginProperty Font 
         Name            =   "IndustryIncW00-Base"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7560
      TabIndex        =   3
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label lblQuote 
      BackStyle       =   0  'Transparent
      Caption         =   "ALL HANDS MAN YOUR BATTLE STATIONS!"
      BeginProperty Font 
         Name            =   "Purista"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   4935
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Programmed by Jeffrey Li 12H. ©Copyrighted RCI ICS4U Games 2016"
      BeginProperty Font 
         Name            =   "IndustryIncW00-Base"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   4440
      Width           =   4695
   End
   Begin VB.Label lblLogo 
      BackStyle       =   0  'Transparent
      Caption         =   "BATTLESHIP"
      BeginProperty Font 
         Name            =   "IndustryIncW00-Base"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5655
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    
    CentreForm Me
    
    
End Sub

Private Sub tmrTimer_Timer()
    
    frmMenu.Show
    Unload Me
    
End Sub
