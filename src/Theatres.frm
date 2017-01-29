VERSION 5.00
Begin VB.Form frmTheatres 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00400000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Battleship"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8790
   Icon            =   "Theatres.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   477
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   586
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optGameplay 
      BackColor       =   &H00400000&
      Caption         =   "GAMEPLAY"
      BeginProperty Font 
         Name            =   "Purista"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1920
      TabIndex        =   12
      Top             =   6480
      Width           =   1455
   End
   Begin VB.OptionButton optStory 
      BackColor       =   &H00400000&
      Caption         =   "SITUATION"
      BeginProperty Font 
         Name            =   "Purista"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   11
      Top             =   6480
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00404040&
      Caption         =   "RETURN"
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
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6360
      Width           =   3375
   End
   Begin VB.Frame fraHard 
      BackColor       =   &H00400000&
      Caption         =   "Northern Command"
      BeginProperty Font 
         Name            =   "IndustryIncW00-Base"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   5295
      Left            =   5880
      TabIndex        =   3
      Top             =   960
      Width           =   2775
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "NOT AVAILABLE IN ALPHA"
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
         Height          =   735
         Left            =   240
         TabIndex        =   20
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Image imgFlag 
         Height          =   375
         Index           =   2
         Left            =   240
         Picture         =   "Theatres.frx":5C12
         Stretch         =   -1  'True
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label lblEnemy 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "NORTHERN FLEET"
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
         Height          =   375
         Index           =   2
         Left            =   840
         TabIndex        =   18
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00400000&
         Caption         =   "HARD SITUATION"
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
         Height          =   1935
         Index           =   2
         Left            =   240
         TabIndex        =   15
         Top             =   3120
         Width           =   2295
      End
      Begin VB.Label lblHard 
         BackStyle       =   0  'Transparent
         Caption         =   "[EXPERT]"
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
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label lblHard_Loc 
         BackStyle       =   0  'Transparent
         Caption         =   "Bering Strait"
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
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   2295
      End
      Begin VB.Image imgMap 
         Enabled         =   0   'False
         Height          =   1575
         Index           =   2
         Left            =   240
         Picture         =   "Theatres.frx":9C16
         Stretch         =   -1  'True
         Top             =   960
         Width           =   2295
      End
   End
   Begin VB.Frame fraMedium 
      BackColor       =   &H00400000&
      Caption         =   "US Seventh Fleet"
      BeginProperty Font 
         Name            =   "IndustryIncW00-Base"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   5295
      Left            =   3000
      TabIndex        =   2
      Top             =   960
      Width           =   2775
      Begin VB.Label lblAlpha 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "NOT AVAILABLE IN ALPHA"
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
         Height          =   735
         Left            =   240
         TabIndex        =   19
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Image imgFlag 
         Height          =   375
         Index           =   1
         Left            =   240
         Picture         =   "Theatres.frx":2D4E2
         Stretch         =   -1  'True
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label lblEnemy 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "SOUTH SEA FLEET"
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
         Height          =   375
         Index           =   1
         Left            =   840
         TabIndex        =   17
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00400000&
         Caption         =   "MEDIUM SITUATION"
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
         Height          =   1935
         Index           =   1
         Left            =   240
         TabIndex        =   14
         Top             =   3120
         Width           =   2295
      End
      Begin VB.Label lblMedium 
         BackStyle       =   0  'Transparent
         Caption         =   "[INTERMEDIATE]"
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
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label lblMedium_Loc 
         BackStyle       =   0  'Transparent
         Caption         =   "South China Sea"
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
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   2295
      End
      Begin VB.Image imgMap 
         Enabled         =   0   'False
         Height          =   1575
         Index           =   1
         Left            =   240
         Picture         =   "Theatres.frx":314E6
         Stretch         =   -1  'True
         Top             =   960
         Width           =   2295
      End
   End
   Begin VB.Frame fraEasy 
      BackColor       =   &H00400000&
      Caption         =   "US Fourth Fleet"
      BeginProperty Font 
         Name            =   "IndustryIncW00-Base"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   5295
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   2775
      Begin VB.Image imgFlag 
         Height          =   375
         Index           =   0
         Left            =   240
         Picture         =   "Theatres.frx":3A73A
         Stretch         =   -1  'True
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label lblEnemy 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "BLACK SEA FLEET"
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
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   16
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00400000&
         Caption         =   "EASY SITUATION"
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
         Height          =   1935
         Index           =   0
         Left            =   240
         TabIndex        =   13
         Top             =   3120
         Width           =   2295
      End
      Begin VB.Label lblEasy 
         BackStyle       =   0  'Transparent
         Caption         =   "[BEGINNER]"
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
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label lblEasy_Loc 
         BackStyle       =   0  'Transparent
         Caption         =   "Black Sea"
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
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   2295
      End
      Begin VB.Image imgMap 
         Height          =   1575
         Index           =   0
         Left            =   240
         Picture         =   "Theatres.frx":3E73E
         Stretch         =   -1  'True
         Top             =   960
         Width           =   2295
      End
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      BackStyle       =   0  'Transparent
      Caption         =   "SELECT YOUR THEATRE"
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
      TabIndex        =   0
      Top             =   120
      Width           =   8535
   End
End
Attribute VB_Name = "frmTheatres"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdReturn_Click()
    
    Unload Me
    
End Sub

Private Sub Form_Load()
    
    SetColours
    DisplaySituations
    CentreForm Me
    
End Sub

Private Sub SetColours()

    Dim K As Integer
    Dim RGBValue As Long
    RGBValue = RGB(40, 60, 70)

    Me.BackColor = RGBValue
    lblTitle.BackColor = RGBValue
    fraEasy.BackColor = RGBValue
    fraMedium.BackColor = RGBValue
    fraHard.BackColor = RGBValue
    optStory.BackColor = RGBValue
    optGameplay.BackColor = RGBValue
    
    For K = 1 To 3
        lblInfo(K - 1).BackColor = RGBValue
    Next K
    
    cmdReturn.BackColor = CMDBOXCOLOR

End Sub

Private Sub DisplaySituations()
    
    Dim Msg As String
    Dim K As Integer
    
    For K = 1 To 3
        lblInfo(K - 1).FontItalic = True
    Next K
    
    Msg = "Novorussian separatists have managed to acquire old Soviet vessels "
    Msg = Msg & "and are threatening the coastlines of neighbouring states. "
    Msg = Msg & "The US Fourth Fleet is mobilized to elimate all hostile vessels in the area. "
    lblInfo(0).Caption = Msg
    
    Msg = "Maritime disputes and escalating tensions have finally broke out into a "
    Msg = Msg & "standoff between the Chinese and US navies. "
    Msg = Msg & "The world holds its breath as the Sino-American War has begun."
    lblInfo(1).Caption = Msg
    
    Msg = "Increasing military presence in the Arctic Ocean has erupted into an armed conflict "
    Msg = Msg & "between Russia and NATO. USNORTHCOM fleets are ordered to disrupt the "
    Msg = Msg & "Russian Navy stationed off the coast of Alaska. "
    lblInfo(2).Caption = Msg
    
End Sub

Private Sub DisplayGameplayInfo()

    Dim Msg As String
    Dim K As Integer
    
    For K = 1 To 3
        lblInfo(K - 1).FontItalic = False
    Next K
    
    Msg = "The opposing navy will fire at random positions on your board. "
    Msg = Msg & "None of your ships will be locked on if they are hit."
    lblInfo(0).Caption = Msg
    
    Msg = "The opposing navy will initially fire at random positions on your board, "
    Msg = Msg & "but the enemy may concentrate their fire on your larger ships."
    lblInfo(1).Caption = Msg
    
    Msg = "The opposing navy will initially fire at positions close to your ships. "
    Msg = Msg & "If any of your ships are hit, the enemy will always concentrate on that ship "
    Msg = Msg & "until it is completely destroyed."
    lblInfo(2).Caption = Msg
    
End Sub

Private Sub imgMap_Click(Index As Integer)
    
    Difficulty = Index
    PassiveClose = True
    
    Unload Me
    If TheatreFromMenu Then
        Unload frmMenu
    Else
        Unload frmGame
    End If
    frmGame.Show
    
End Sub

Private Sub imgMap_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    imgMap(Index).BorderStyle = 1
    
End Sub

Private Sub imgMap_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    imgMap(Index).BorderStyle = 0
    
End Sub

Private Sub optGameplay_Click()
    
    DisplayGameplayInfo
    
End Sub

Private Sub optStory_Click()
    
    DisplaySituations
    
End Sub
