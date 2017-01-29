VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmGame 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Battleship 2016 - 2.0 Alpha Build"
   ClientHeight    =   8775
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   15030
   Icon            =   "Game.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   585
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1002
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrGame 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   240
      Top             =   120
   End
   Begin VB.Frame fraDialog 
      BackColor       =   &H00FF8080&
      Caption         =   "Placements"
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
      Height          =   975
      Left            =   2040
      TabIndex        =   28
      Top             =   6960
      Width           =   5415
      Begin VB.OptionButton optHorz 
         BackColor       =   &H00FF8080&
         Caption         =   "Horizontal"
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
         Height          =   195
         Left            =   840
         TabIndex        =   30
         Top             =   480
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton optVert 
         BackColor       =   &H00FF8080&
         Caption         =   "Vertical"
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
         Height          =   195
         Left            =   2760
         TabIndex        =   29
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.Frame fraStats 
      BackColor       =   &H00FF8080&
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
      Height          =   975
      Left            =   7680
      TabIndex        =   19
      Top             =   6960
      Width           =   5415
      Begin VB.Label lblAcc 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "00.00%"
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
         Left            =   1920
         TabIndex        =   27
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblTurn 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Left            =   3840
         TabIndex        =   26
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lblMisses 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Left            =   3840
         TabIndex        =   25
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblHits 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Left            =   1920
         TabIndex        =   24
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblLTurn 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Turn:"
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
         Left            =   2880
         TabIndex        =   23
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblLMisses 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Misses:"
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
         Left            =   2880
         TabIndex        =   22
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblLHits 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Hits:"
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
         Left            =   840
         TabIndex        =   21
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblLAccuracy 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Accuracy:"
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
         Left            =   840
         TabIndex        =   20
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H00404040&
      Caption         =   "BEGIN OPERATION"
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   8040
      Width           =   3495
   End
   Begin VB.CommandButton cmdRandom 
      BackColor       =   &H00404040&
      Caption         =   "RANDOM PLACEMENT"
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
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   8040
      Width           =   3495
   End
   Begin VB.CommandButton cmdReset 
      BackColor       =   &H00404040&
      Caption         =   "RESET POSITIONS"
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
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   8040
      Width           =   3495
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
      Left            =   11400
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   8040
      Width           =   3495
   End
   Begin MSFlexGridLib.MSFlexGrid grdComputer 
      Height          =   5370
      Left            =   7800
      TabIndex        =   1
      Top             =   1080
      Width           =   5370
      _ExtentX        =   9472
      _ExtentY        =   9472
      _Version        =   393216
      Rows            =   11
      Cols            =   11
      RowHeightMin    =   480
      BackColor       =   4194304
      FocusRect       =   0
      HighLight       =   0
      ScrollBars      =   0
   End
   Begin MSFlexGridLib.MSFlexGrid grdPlayer 
      Height          =   5370
      Left            =   2040
      TabIndex        =   0
      Top             =   1080
      Width           =   5370
      _ExtentX        =   9472
      _ExtentY        =   9472
      _Version        =   393216
      Rows            =   11
      Cols            =   11
      RowHeightMin    =   480
      BackColor       =   4194304
      Enabled         =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      ScrollBars      =   0
   End
   Begin VB.Image imgColour 
      Height          =   480
      Index           =   2
      Left            =   15360
      Picture         =   "Game.frx":5C12
      Top             =   4920
      Width           =   480
   End
   Begin VB.Image imgColour 
      Height          =   480
      Index           =   1
      Left            =   15360
      Picture         =   "Game.frx":6854
      Top             =   4320
      Width           =   480
   End
   Begin VB.Image imgColour 
      Height          =   480
      Index           =   0
      Left            =   15360
      Picture         =   "Game.frx":7496
      Top             =   3720
      Width           =   480
   End
   Begin VB.Label lblDialog 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "INSERT DIALOG HERE!"
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
      Left            =   2040
      TabIndex        =   31
      Top             =   6600
      Width           =   11055
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "00:00"
      BeginProperty Font 
         Name            =   "Purista"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   7080
      TabIndex        =   15
      Top             =   360
      Width           =   975
   End
   Begin VB.Image imgEFlag 
      Height          =   375
      Index           =   2
      Left            =   15240
      Picture         =   "Game.frx":80D8
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   615
   End
   Begin VB.Image imgEFlag 
      Height          =   375
      Index           =   1
      Left            =   15240
      Picture         =   "Game.frx":C0DC
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   615
   End
   Begin VB.Image imgEFlag 
      Height          =   375
      Index           =   0
      Left            =   15240
      Picture         =   "Game.frx":100E0
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label lblCShip 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CORVETTE"
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
      Left            =   13320
      TabIndex        =   13
      Top             =   7560
      Width           =   1575
   End
   Begin VB.Label lblCShip 
      Alignment       =   1  'Right Justify
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
      Left            =   13320
      TabIndex        =   12
      Top             =   6120
      Width           =   1575
   End
   Begin VB.Label lblCShip 
      Alignment       =   1  'Right Justify
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
      Left            =   13320
      TabIndex        =   11
      Top             =   4680
      Width           =   1575
   End
   Begin VB.Label lblCShip 
      Alignment       =   1  'Right Justify
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
      Left            =   13320
      TabIndex        =   10
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label lblCShip 
      Alignment       =   1  'Right Justify
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
      Left            =   13320
      TabIndex        =   9
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label lblPShip 
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
      Left            =   240
      TabIndex        =   8
      Top             =   7560
      Width           =   1575
   End
   Begin VB.Label lblPShip 
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
      Left            =   240
      TabIndex        =   7
      Top             =   6120
      Width           =   1575
   End
   Begin VB.Label lblPShip 
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
      Left            =   240
      TabIndex        =   6
      Top             =   4680
      Width           =   1575
   End
   Begin VB.Label lblPShip 
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
      Left            =   240
      TabIndex        =   5
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label lblPShip 
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
      Left            =   240
      TabIndex        =   4
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Image imgFlag 
      Height          =   540
      Index           =   1
      Left            =   12240
      Picture         =   "Game.frx":140E4
      Stretch         =   -1  'True
      Top             =   360
      Width           =   855
   End
   Begin VB.Image imgFlag 
      Height          =   540
      Index           =   0
      Left            =   2040
      Picture         =   "Game.frx":180E8
      Stretch         =   -1  'True
      Top             =   360
      Width           =   855
   End
   Begin VB.Label lblNavy 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "OPFOR NAVY"
      BeginProperty Font 
         Name            =   "IndustryIncW00-Base"
         Size            =   21
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   1
      Left            =   8280
      TabIndex        =   3
      Top             =   360
      Width           =   3855
   End
   Begin VB.Label lblNavy 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "PLAYER NAVY"
      BeginProperty Font 
         Name            =   "IndustryIncW00-Base"
         Size            =   21
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   0
      Left            =   3000
      TabIndex        =   2
      Top             =   360
      Width           =   3975
   End
   Begin VB.Image imgCShip 
      Height          =   1095
      Index           =   4
      Left            =   13320
      Picture         =   "Game.frx":1C0EC
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   1575
   End
   Begin VB.Image imgCShip 
      Height          =   1095
      Index           =   3
      Left            =   13320
      Picture         =   "Game.frx":2BE50
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Image imgCShip 
      Height          =   1095
      Index           =   2
      Left            =   13320
      Picture         =   "Game.frx":3BBB4
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Image imgCShip 
      Height          =   1095
      Index           =   1
      Left            =   13320
      Picture         =   "Game.frx":4B918
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Image imgCShip 
      Height          =   1095
      Index           =   0
      Left            =   13320
      Picture         =   "Game.frx":5B67C
      Stretch         =   -1  'True
      Top             =   720
      Width           =   1575
   End
   Begin VB.Image imgPShip 
      Height          =   1095
      Index           =   4
      Left            =   240
      Picture         =   "Game.frx":6B3E0
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   1575
   End
   Begin VB.Image imgPShip 
      Height          =   1095
      Index           =   3
      Left            =   240
      Picture         =   "Game.frx":7B144
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Image imgPShip 
      Height          =   1095
      Index           =   2
      Left            =   240
      Picture         =   "Game.frx":8AEA8
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Image imgPShip 
      Height          =   1095
      Index           =   1
      Left            =   240
      Picture         =   "Game.frx":9AC0C
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Image imgPShip 
      Height          =   1095
      Index           =   0
      Left            =   240
      Picture         =   "Game.frx":AA970
      Stretch         =   -1  'True
      Top             =   720
      Width           =   1575
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   504
      X2              =   504
      Y1              =   72
      Y2              =   424
   End
   Begin VB.Image imgGPShip 
      Height          =   1095
      Index           =   4
      Left            =   240
      Picture         =   "Game.frx":BA6D4
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   1575
   End
   Begin VB.Image imgGPShip 
      Height          =   1095
      Index           =   3
      Left            =   240
      Picture         =   "Game.frx":CA438
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Image imgGPShip 
      Height          =   1095
      Index           =   2
      Left            =   240
      Picture         =   "Game.frx":DA19C
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Image imgGPShip 
      Height          =   1095
      Index           =   1
      Left            =   240
      Picture         =   "Game.frx":E9F00
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Image imgGPShip 
      Height          =   1095
      Index           =   0
      Left            =   240
      Picture         =   "Game.frx":F9C64
      Stretch         =   -1  'True
      Top             =   720
      Width           =   1575
   End
   Begin VB.Image imgGCShip 
      Height          =   1095
      Index           =   4
      Left            =   13320
      Picture         =   "Game.frx":1099C8
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   1575
   End
   Begin VB.Image imgGCShip 
      Height          =   1095
      Index           =   3
      Left            =   13320
      Picture         =   "Game.frx":11972C
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Image imgGCShip 
      Height          =   1095
      Index           =   2
      Left            =   13320
      Picture         =   "Game.frx":129490
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Image imgGCShip 
      Height          =   1095
      Index           =   1
      Left            =   13320
      Picture         =   "Game.frx":1391F4
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Image imgGCShip 
      Height          =   1095
      Index           =   0
      Left            =   13320
      Picture         =   "Game.frx":148F58
      Stretch         =   -1  'True
      Top             =   720
      Width           =   1575
   End
   Begin VB.Menu mnuGame 
      Caption         =   "Game"
      Begin VB.Menu mnuNewGame 
         Caption         =   "New Game"
      End
      Begin VB.Menu mnuLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReturn 
         Caption         =   "Exit to Menu"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit to Desktop"
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
   End
End
Attribute VB_Name = "frmGame"
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

Const FORM_WIDTH = 15120
Const FORM_HEIGHT = 9495
Const UNSELECTED = -1
Const MAXSHIPS = 5

Dim PassiveCloseGame As Boolean

Dim PlacedShips As Integer
Dim SelectedClass As Integer
Dim Direction As Integer
Dim GameStart As Boolean
Dim GameOver As Boolean
Dim PSunks, CSunks As Integer

Option Explicit

Private Sub InitializeFormValues()
    
    Dim K As Integer
    
    GameStart = False
    GameOver = False
    Victorious = False
    
    Hits = 0
    Misses = 0
    Accuracy = 0
    Turn = 0
    PSunks = 0
    CSunks = 0
    GameTime = 0
    PlacedShips = 0
    SelectedClass = UNSELECTED
    Direction = 0
    
    optHorz.Value = True
    optVert.Value = False
    
    For K = 0 To 4
        imgPShip(K).Enabled = True
        imgGPShip(K).Enabled = True
    Next K
    grdPlayer.Enabled = True
    grdComputer.Enabled = True
    
    lblHits.Caption = Hits
    lblMisses.Caption = Misses
    lblAcc.Caption = "00.00%"
    lblTurn.Caption = Turn

End Sub

Private Sub cmdRandom_Click()
    
    Dim K As Integer
    
    For K = 0 To 4
        imgPShip(K).Visible = False
        lblPShip(K).ForeColor = vbGreen
    Next K
    
    With grdPlayer
        .Visible = False
        RandomPlaceShips grdPlayer, PCellValue()
        .Visible = True
    End With
    PlacedShips = MAXSHIPS
    
    cmdRandom.Enabled = False
    cmdStart.Enabled = True
    cmdReset.Enabled = True
    optHorz.Enabled = False
    optVert.Enabled = False
    
    lblDialog.Caption = UCase$(GetDialogMsg(4))
    
End Sub

Private Sub cmdReset_Click()
    
    Dim K As Integer
    
    tmrGame.Enabled = False
    lblTime.Caption = "00:00"
    GameTime = 0
    cmdReset.Enabled = False
    cmdStart.Enabled = False
    cmdRandom.Enabled = True
    optHorz.Enabled = True
    optVert.Enabled = True
    
    With grdPlayer
    
        .Visible = False
        
        PlacedShips = 0
        SelectedClass = UNSELECTED
        
        For K = 0 To 4
            imgPShip(K).Visible = True
            lblPShip(K).ForeColor = vbWhite
        Next K
        
        InitializeArrayValues PCellValue()
        InitializeBoard grdPlayer
        
        .Visible = True
    
    End With
    
    lblDialog.Caption = UCase$(GetDialogMsg(0))
    
    If GameStart Then
        
        InitializeFormValues
        
        cmdReset.Caption = "RESET POSITIONS"
        
        With grdComputer
    
            .Visible = False
            
            For K = 0 To 4
                imgCShip(K).Enabled = True
                imgCShip(K).Visible = True
                lblCShip(K).ForeColor = vbWhite
            Next K
            
            InitializeArrayValues CCellValue()
            InitializeBoard grdComputer
            
            .Visible = True
    
        End With
        
        GameStart = False
        
    End If
    
End Sub

Private Sub cmdReturn_Click()
    
    mnuReturn_Click
    
End Sub

Private Sub SetColours()
    
    Me.BackColor = RGB(40, 60, 70)
    
    cmdStart.BackColor = CMDBOXCOLOR
    cmdRandom.BackColor = CMDBOXCOLOR
    cmdReset.BackColor = CMDBOXCOLOR
    cmdReturn.BackColor = CMDBOXCOLOR
    
    fraDialog.BackColor = RGB(40, 60, 70)
    fraStats.BackColor = RGB(40, 60, 70)
    optHorz.BackColor = RGB(40, 60, 70)
    optVert.BackColor = RGB(40, 60, 70)

End Sub

Private Sub cmdStart_Click()
      
    Dim K As Integer
    Dim CoinFlip As Integer
    Dim Msg As String
    Dim CResult As String
    
    For K = 0 To 4
        imgPShip(K).Visible = True
        lblPShip(K).ForeColor = vbWhite
    Next K
    
    cmdRandom.Enabled = False
    cmdStart.Enabled = False
    cmdReset.Caption = "RESTART GAME"
    
    With grdComputer
        .Visible = False
        RandomPlaceShips grdComputer, CCellValue(), False
        .Visible = True
    End With
    
    CoinFlip = GetRandom(0, 1)
    
    Turn = 0
    lblTurn.Caption = 1
    If CoinFlip = 0 Then
        lblDialog.Caption = UCase$(GetDialogMsg(5))
    Else
        lblDialog.Caption = UCase$(GetDialogMsg(6))
        CResult = ComputerFireTurnResult(grdPlayer, Msg, lblPShip, imgPShip, imgColour, Difficulty)
    End If
    
    GameStart = True
    tmrGame.Enabled = True
    
End Sub

Private Sub Form_Click()
    
    Dim K As Integer
    Dim Msg As String
    
    If SelectedClass <> UNSELECTED Then
        If PlacedShips = MAXSHIPS Then
            lblDialog.Caption = UCase$(GetDialogMsg(4))
        Else
            imgPShip(SelectedClass).BorderStyle = 0
            lblPShip(SelectedClass).ForeColor = vbWhite
            SelectedClass = UNSELECTED
            lblDialog.Caption = UCase$(GetDialogMsg(0))
        End If
    End If
    
End Sub

Private Sub Form_Load()
    
    Dim K As Integer
    Dim Msg As String
    
    PassiveCloseGame = False
    
    InitializeFormValues
    InitializeBoard grdPlayer
    InitializeBoard grdComputer
    InitializeArrayValues PCellValue()
    InitializeArrayValues CCellValue()
    SetColours
    SetGameFormControls Difficulty, lblNavy, lblPShip, lblCShip, imgEFlag, imgFlag
    
    lblDialog.Caption = UCase$(GetDialogMsg(0))
    
    Me.Width = FORM_WIDTH
    Me.Height = FORM_HEIGHT
    CentreForm Me
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
            
'    If Not PassiveCloseGame Then
'        If UnloadMode = 0 Then
'            Cancel = 1
'        End If
'        mnuExit_Click
'    End If
    
End Sub

Private Sub grdComputer_Click()
    
    Dim K As Integer
    Dim PMsg As String
    Dim CMsg As String
    Dim PResult, CResult As Integer
    Dim ValidHit As Boolean
    
    With grdComputer
    
    If GameStart Then
        
        If PSunks < 5 And CSunks < 5 Then
            
            .Visible = False
            
            PResult = PlayerFireTurnResult(grdComputer, PMsg, Hits, Misses, _
                lblCShip, imgCShip, imgColour)
            lblDialog.Caption = UCase$(PMsg)
            If PResult = SUNK Then
                PSunks = PSunks + 1
            End If
            
            If PResult <> REHIT Then
                Turn = Turn + 1
                lblTurn.Caption = Turn
                
                CResult = ComputerFireTurnResult(grdPlayer, CMsg, lblPShip, imgPShip, _
                    imgColour, Difficulty)
                lblDialog.Caption = UCase$(PMsg & " " & CMsg)
                If CResult = SUNK Then
                    CSunks = CSunks + 1
                End If
            End If
            
            lblHits.Caption = Hits
            lblMisses.Caption = Misses
            Accuracy = Hits / (Hits + Misses)
            lblAcc.Caption = Format$(Accuracy, "00.00%")
        
            .Visible = True
        
        End If
        
        If PSunks = 5 Or CSunks = 5 Then
            GameOver = True
            tmrGame.Enabled = False
            If PSunks = 5 Then
                Victorious = True
                lblDialog.Caption = UCase$(GetDialogMsg(7))
            ElseIf CSunks = 5 Then
                lblDialog.Caption = UCase$(GetDialogMsg(8))
            End If
            For K = 0 To 4
                imgPShip(K).Enabled = False
                imgGPShip(K).Enabled = False
            Next K
            grdComputer.Enabled = False
            grdPlayer.Enabled = False
            frmGameOver.Show vbModal
        End If
        
    End If
    
    End With
    
End Sub

Private Sub grdPlayer_Click()
        
    Dim Msg As String
    Dim ORow, OCol As Integer
    Dim ValidPlace As Boolean
    
    With grdPlayer
    
    ORow = .Row
    OCol = .Col
    
    .Visible = False
    
    If SelectedClass <> UNSELECTED Then
    
        imgPShip(SelectedClass).BorderStyle = 0
        lblPShip(SelectedClass).ForeColor = vbWhite
        CheckAndPlaceShip grdPlayer, PCellValue(), SelectedClass, Direction, ValidPlace
        
        If Not ValidPlace Then
            Msg = "Invalid position. Selected ship overlaps with another ship."
        Else
            PlacedShips = PlacedShips + 1
            imgPShip(SelectedClass).Visible = False
            lblPShip(SelectedClass).ForeColor = vbGreen
            cmdReset.Enabled = True
            cmdRandom.Enabled = False
            If PlacedShips = 5 Then
                cmdStart.Enabled = True
                optHorz.Enabled = False
                optVert.Enabled = False
                Msg = GetDialogMsg(4)
            Else
                Msg = GetDialogMsg(0)
            End If
        End If
        
        lblDialog.Caption = UCase$(Msg)
        
        SelectedClass = UNSELECTED
        
    ElseIf GameStart And Not GameOver Then
    
        Msg = "Uhh... sir? Do you want us to fire at our own ships?"
        lblDialog.Caption = UCase$(Msg)
        
    End If
    
    .Row = ORow
    .Col = OCol
    
    .Visible = True
    
    End With
    
End Sub

Private Sub imgGPShip_Click(Index As Integer)
    
    Dim Msg As String
    
    If GameStart Then
        Msg = "Our " & GetShipClass(Index) & " has been sunken."
    End If
    
End Sub

Private Sub imgPShip_Click(Index As Integer)
    
    Dim Msg As String
    Dim ShipType As Integer
    Dim ShipValue, ShipLength As Integer
    Dim CCol, CRow As Integer
    Dim HP, Dmg As Integer
    
    With grdPlayer
    
    If GameStart Then
        ShipValue = Index + 1
        ShipLength = GetShipLength(Index)
        HP = 0
        For CRow = 1 To .Rows - 1
            For CCol = 1 To .Cols - 1
                If PCellValue(CCol, CRow) = ShipValue Then
                    HP = HP + 1
                End If
            Next CCol
        Next CRow
        
        Dmg = ShipLength - HP
        If Dmg = 0 Then
            Msg = "Our " & GetShipClass(Index) & " is fully operational."
        Else
            Msg = "Our " & GetShipClass(Index) & " has been hit. "
            Msg = Msg & "Damage is at " & Format$(Dmg / ShipLength, "00.00%") & "."
        End If
        lblDialog.Caption = UCase$(Msg)
    Else
    
        If SelectedClass <> UNSELECTED Then
            imgPShip(SelectedClass).BorderStyle = 0
            lblPShip(SelectedClass).ForeColor = vbWhite
        End If
        
        SelectedClass = Index
        imgPShip(Index).BorderStyle = 1
        lblPShip(Index).ForeColor = vbYellow
        
        Msg = GetShipClass(Index) & " selected. [" & GetShipLength(Index) & " units long]"
        lblDialog.Caption = UCase$(Msg)
    
    End If
    
    End With
    
End Sub

Private Sub lblDialog_Click()
    
    Form_Click
    
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
     
    PassiveCloseGame = True
    TheatreFromMenu = False
    frmTheatres.Show vbModal
    
End Sub

Private Sub mnuReturn_Click()
        
    PassiveCloseGame = True
    If GameStart Then
        frmReturn.Show vbModal
    Else
        Unload Me
        frmMenu.Show
    End If
    
End Sub

Private Sub optHorz_Click()
    
    Direction = 0
    
End Sub

Private Sub optVert_Click()
    
    Direction = 1
    
End Sub

Private Sub tmrGame_Timer()
    
    GameTime = GameTime + 1
    lblTime.Caption = Format$(GameTime \ 60, "00") & ":" & _
        Format$(GameTime Mod 60, "00")
        
End Sub
