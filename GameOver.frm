VERSION 5.00
Begin VB.Form frmGameOver 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00400000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Game Over"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9510
   Icon            =   "GameOver.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   9510
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraStats 
      BackColor       =   &H00400000&
      Caption         =   "Statistics"
      BeginProperty Font 
         Name            =   "IndustryIncW00-Base"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   3840
      TabIndex        =   5
      Top             =   3000
      Width           =   5535
      Begin VB.Label lblTime 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "00:00"
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
         Height          =   360
         Left            =   3000
         TabIndex        =   15
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblLTime 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Game Time:"
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
         Height          =   360
         Left            =   1440
         TabIndex        =   14
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblLAccuracy 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Accuracy:"
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
         Height          =   360
         Left            =   480
         TabIndex        =   13
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label lblLHits 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Hits:"
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
         Height          =   360
         Left            =   480
         TabIndex        =   12
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label lblLMisses 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Misses:"
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
         Height          =   360
         Left            =   3120
         TabIndex        =   11
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblLTurn 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Turns:"
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
         Height          =   360
         Left            =   3120
         TabIndex        =   10
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label lblHits 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Height          =   360
         Left            =   1800
         TabIndex        =   9
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label lblMisses 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Height          =   360
         Left            =   4320
         TabIndex        =   8
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lblTurn 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Height          =   360
         Left            =   4320
         TabIndex        =   7
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label lblAcc 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "00.00%"
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
         Height          =   360
         Left            =   1800
         TabIndex        =   6
         Top             =   840
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdMenu 
      BackColor       =   &H00404040&
      Caption         =   "RETURN TO MENU"
      BeginProperty Font 
         Name            =   "IndustryIncW00-Base"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      MaskColor       =   &H00FF8080&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5160
      Width           =   3015
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00404040&
      Caption         =   "RETURN TO GAME"
      BeginProperty Font 
         Name            =   "IndustryIncW00-Base"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      MaskColor       =   &H00FF8080&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5160
      Width           =   3015
   End
   Begin VB.CommandButton cmdNewGame 
      BackColor       =   &H00404040&
      Caption         =   "NEW GAME"
      BeginProperty Font 
         Name            =   "IndustryIncW00-Base"
         Size            =   14.25
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
      TabIndex        =   1
      Top             =   5160
      Width           =   3015
   End
   Begin VB.Label lblDialog 
      BackStyle       =   0  'Transparent
      Caption         =   "INSERT DIALOG HERE!"
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
      Height          =   1575
      Left            =   3840
      TabIndex        =   4
      Top             =   1200
      Width           =   5415
   End
   Begin VB.Image imgVictory 
      Height          =   4890
      Left            =   120
      Picture         =   "GameOver.frx":5C12
      Stretch         =   -1  'True
      Top             =   120
      Width           =   3420
   End
   Begin VB.Label lblResult 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      BackStyle       =   0  'Transparent
      Caption         =   "VICTORY"
      BeginProperty Font 
         Name            =   "IndustryIncW00-Base"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   3720
      TabIndex        =   0
      Top             =   120
      Width           =   5655
   End
   Begin VB.Image imgDefeat 
      Height          =   4890
      Left            =   120
      Picture         =   "GameOver.frx":18C1B
      Stretch         =   -1  'True
      Top             =   120
      Width           =   3420
   End
End
Attribute VB_Name = "frmGameOver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdMenu_Click()
    
    Unload Me
    Unload frmGame
    frmMenu.Show
    
End Sub

Private Sub cmdNewGame_Click()
    
    PassiveCloseGame = True
    Unload Me
    Unload frmGame
    frmTheatres.Show vbModal
    
End Sub

Private Sub cmdReturn_Click()
    
    Unload Me
    
End Sub

Private Sub Form_Load()
    
    Dim Msg As String
    
    Me.BackColor = RGB(40, 60, 70)
    fraStats.BackColor = Me.BackColor
    cmdNewGame.BackColor = CMDBOXCOLOR
    cmdReturn.BackColor = CMDBOXCOLOR
    cmdMenu.BackColor = CMDBOXCOLOR
    
    lblHits.Caption = Hits
    lblMisses.Caption = Misses
    lblAcc.Caption = Format$(Accuracy, "00.00%")
    lblTurn.Caption = Turn
    lblTime.Caption = Format$(GameTime \ 60, "00") & ":" & _
        Format$(GameTime Mod 60, "00")
    
    If Victorious Then
        lblResult.Caption = "VICTORY"
        imgVictory.Visible = True
        imgDefeat.Visible = False
        Msg = "We have successfully eliminated all hostile ships within this region. "
        Msg = Msg & "The enemy fleet is retreating, but the war is not over yet. "
        Msg = Msg & "We need new leaders and you, Captain, have just proven yourself. "
        Msg = Msg & "Congratulations, Admiral!"
        lblDialog.Caption = Msg
    Else
        lblResult.Caption = "DEFEAT"
        imgVictory.Visible = False
        imgDefeat.Visible = True
        Msg = "It seems this time the enemy has managed to destroy our battlegroup within "
        Msg = Msg & "this region. Our seapower has been challenged. We are assessing all damages "
        Msg = Msg & "and casualties. The battle maybe lost, but the war has just begun."
        lblDialog.Caption = Msg
    End If
    
    CentreForm Me
    
End Sub
