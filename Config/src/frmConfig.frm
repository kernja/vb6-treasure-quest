VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmConfig 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Treasure Quest Configuration"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   4425
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab tabMenu 
      Height          =   3855
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   6800
      _Version        =   393216
      Style           =   1
      Tab             =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Profiles"
      TabPicture(0)   =   "frmConfig.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblInfo(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtProfileName"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "btnProfile(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lstProfiles"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "btnProfile(1)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Audio / Video Configuration"
      TabPicture(1)   =   "frmConfig.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lblInfo(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fmeOption(0)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "fmeOption(1)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Controls"
      TabPicture(2)   =   "frmConfig.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "tabInput"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "lblInfo(3)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      Begin TabDlg.SSTab tabInput 
         Height          =   2775
         Left            =   -74760
         TabIndex        =   17
         Top             =   840
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   4895
         _Version        =   393216
         Tabs            =   4
         TabsPerRow      =   4
         TabHeight       =   520
         TabCaption(0)   =   "Player 1"
         TabPicture(0)   =   "frmConfig.frx":0054
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "optInput(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "optInput(1)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "fmeOption(2)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "btnInput(0)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).ControlCount=   4
         TabCaption(1)   =   "Player 2"
         TabPicture(1)   =   "frmConfig.frx":0070
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "btnInput(1)"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "fmeOption(3)"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "optInput(3)"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "optInput(2)"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).ControlCount=   4
         TabCaption(2)   =   "Player 3"
         TabPicture(2)   =   "frmConfig.frx":008C
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "btnInput(2)"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "fmeOption(4)"
         Tab(2).Control(1).Enabled=   0   'False
         Tab(2).Control(2)=   "optInput(5)"
         Tab(2).Control(2).Enabled=   0   'False
         Tab(2).Control(3)=   "optInput(4)"
         Tab(2).Control(3).Enabled=   0   'False
         Tab(2).ControlCount=   4
         TabCaption(3)   =   "Player 4"
         TabPicture(3)   =   "frmConfig.frx":00A8
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "btnInput(3)"
         Tab(3).Control(0).Enabled=   0   'False
         Tab(3).Control(1)=   "fmeOption(5)"
         Tab(3).Control(1).Enabled=   0   'False
         Tab(3).Control(2)=   "optInput(7)"
         Tab(3).Control(2).Enabled=   0   'False
         Tab(3).Control(3)=   "optInput(6)"
         Tab(3).Control(3).Enabled=   0   'False
         Tab(3).ControlCount=   4
         Begin VB.CommandButton btnInput 
            Caption         =   "Redefine controls"
            Height          =   255
            Index           =   3
            Left            =   -73080
            TabIndex        =   61
            Top             =   2280
            Width           =   1575
         End
         Begin VB.Frame fmeOption 
            Caption         =   "Current settings"
            Height          =   1335
            Index           =   5
            Left            =   -74760
            TabIndex        =   53
            Top             =   840
            Width           =   3255
            Begin VB.Label lblInfo 
               Caption         =   "Up:"
               Height          =   255
               Index           =   31
               Left            =   360
               TabIndex        =   60
               Top             =   240
               Width           =   1215
            End
            Begin VB.Label lblInfo 
               Caption         =   "Down:"
               Height          =   255
               Index           =   30
               Left            =   360
               TabIndex        =   59
               Top             =   480
               Width           =   1215
            End
            Begin VB.Label lblInfo 
               Caption         =   "Left:"
               Height          =   255
               Index           =   29
               Left            =   360
               TabIndex        =   58
               Top             =   720
               Width           =   1215
            End
            Begin VB.Label lblInfo 
               Caption         =   "Right:"
               Height          =   255
               Index           =   28
               Left            =   360
               TabIndex        =   57
               Top             =   960
               Width           =   1215
            End
            Begin VB.Label lblInfo 
               Caption         =   "Jump:"
               Height          =   255
               Index           =   27
               Left            =   1680
               TabIndex        =   56
               Top             =   240
               Width           =   1455
            End
            Begin VB.Label lblInfo 
               Caption         =   "Fly:"
               Height          =   255
               Index           =   26
               Left            =   1680
               TabIndex        =   55
               Top             =   480
               Width           =   1455
            End
            Begin VB.Label lblInfo 
               Caption         =   "Pause/Select:"
               Height          =   255
               Index           =   25
               Left            =   1680
               TabIndex        =   54
               Top             =   720
               Width           =   1455
            End
         End
         Begin VB.OptionButton optInput 
            Caption         =   "Gamepad / Joystick"
            Height          =   255
            Index           =   7
            Left            =   -73440
            TabIndex        =   52
            Top             =   480
            Width           =   1935
         End
         Begin VB.OptionButton optInput 
            Caption         =   "Keyboard"
            Height          =   255
            Index           =   6
            Left            =   -74760
            TabIndex        =   51
            Top             =   480
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.CommandButton btnInput 
            Caption         =   "Redefine controls"
            Height          =   255
            Index           =   2
            Left            =   -73080
            TabIndex        =   50
            Top             =   2280
            Width           =   1575
         End
         Begin VB.Frame fmeOption 
            Caption         =   "Current settings"
            Height          =   1335
            Index           =   4
            Left            =   -74760
            TabIndex        =   42
            Top             =   840
            Width           =   3255
            Begin VB.Label lblInfo 
               Caption         =   "Up:"
               Height          =   255
               Index           =   24
               Left            =   360
               TabIndex        =   49
               Top             =   240
               Width           =   1215
            End
            Begin VB.Label lblInfo 
               Caption         =   "Down:"
               Height          =   255
               Index           =   23
               Left            =   360
               TabIndex        =   48
               Top             =   480
               Width           =   1215
            End
            Begin VB.Label lblInfo 
               Caption         =   "Left:"
               Height          =   255
               Index           =   22
               Left            =   360
               TabIndex        =   47
               Top             =   720
               Width           =   1215
            End
            Begin VB.Label lblInfo 
               Caption         =   "Right:"
               Height          =   255
               Index           =   21
               Left            =   360
               TabIndex        =   46
               Top             =   960
               Width           =   1215
            End
            Begin VB.Label lblInfo 
               Caption         =   "Jump:"
               Height          =   255
               Index           =   20
               Left            =   1680
               TabIndex        =   45
               Top             =   240
               Width           =   1455
            End
            Begin VB.Label lblInfo 
               Caption         =   "Fly:"
               Height          =   255
               Index           =   19
               Left            =   1680
               TabIndex        =   44
               Top             =   480
               Width           =   1455
            End
            Begin VB.Label lblInfo 
               Caption         =   "Pause/Select:"
               Height          =   255
               Index           =   18
               Left            =   1680
               TabIndex        =   43
               Top             =   720
               Width           =   1455
            End
         End
         Begin VB.OptionButton optInput 
            Caption         =   "Gamepad / Joystick"
            Height          =   255
            Index           =   5
            Left            =   -73440
            TabIndex        =   41
            Top             =   480
            Width           =   1935
         End
         Begin VB.OptionButton optInput 
            Caption         =   "Keyboard"
            Height          =   255
            Index           =   4
            Left            =   -74760
            TabIndex        =   40
            Top             =   480
            Width           =   1695
         End
         Begin VB.CommandButton btnInput 
            Caption         =   "Redefine controls"
            Height          =   255
            Index           =   1
            Left            =   -73080
            TabIndex        =   39
            Top             =   2280
            Width           =   1575
         End
         Begin VB.Frame fmeOption 
            Caption         =   "Current settings"
            Height          =   1335
            Index           =   3
            Left            =   -74760
            TabIndex        =   31
            Top             =   840
            Width           =   3255
            Begin VB.Label lblInfo 
               Caption         =   "Up:"
               Height          =   255
               Index           =   17
               Left            =   360
               TabIndex        =   38
               Top             =   240
               Width           =   1215
            End
            Begin VB.Label lblInfo 
               Caption         =   "Down:"
               Height          =   255
               Index           =   16
               Left            =   360
               TabIndex        =   37
               Top             =   480
               Width           =   1215
            End
            Begin VB.Label lblInfo 
               Caption         =   "Left:"
               Height          =   255
               Index           =   15
               Left            =   360
               TabIndex        =   36
               Top             =   720
               Width           =   1215
            End
            Begin VB.Label lblInfo 
               Caption         =   "Right:"
               Height          =   255
               Index           =   14
               Left            =   360
               TabIndex        =   35
               Top             =   960
               Width           =   1215
            End
            Begin VB.Label lblInfo 
               Caption         =   "Jump:"
               Height          =   255
               Index           =   13
               Left            =   1680
               TabIndex        =   34
               Top             =   240
               Width           =   1455
            End
            Begin VB.Label lblInfo 
               Caption         =   "Fly:"
               Height          =   255
               Index           =   12
               Left            =   1680
               TabIndex        =   33
               Top             =   480
               Width           =   1455
            End
            Begin VB.Label lblInfo 
               Caption         =   "Pause/Select:"
               Height          =   255
               Index           =   11
               Left            =   1680
               TabIndex        =   32
               Top             =   720
               Width           =   1455
            End
         End
         Begin VB.OptionButton optInput 
            Caption         =   "Gamepad / Joystick"
            Height          =   255
            Index           =   3
            Left            =   -73440
            TabIndex        =   30
            Top             =   480
            Width           =   1935
         End
         Begin VB.OptionButton optInput 
            Caption         =   "Keyboard"
            Height          =   255
            Index           =   2
            Left            =   -74760
            TabIndex        =   29
            Top             =   480
            Width           =   1695
         End
         Begin VB.CommandButton btnInput 
            Caption         =   "Redefine controls"
            Height          =   255
            Index           =   0
            Left            =   1920
            TabIndex        =   28
            Top             =   2280
            Width           =   1575
         End
         Begin VB.Frame fmeOption 
            Caption         =   "Current settings"
            Height          =   1335
            Index           =   2
            Left            =   240
            TabIndex        =   20
            Top             =   840
            Width           =   3255
            Begin VB.Label lblInfo 
               Caption         =   "Pause/Select:"
               Height          =   255
               Index           =   10
               Left            =   1680
               TabIndex        =   27
               Top             =   720
               Width           =   1455
            End
            Begin VB.Label lblInfo 
               Caption         =   "Fly:"
               Height          =   255
               Index           =   9
               Left            =   1680
               TabIndex        =   26
               Top             =   480
               Width           =   1455
            End
            Begin VB.Label lblInfo 
               Caption         =   "Jump:"
               Height          =   255
               Index           =   8
               Left            =   1680
               TabIndex        =   25
               Top             =   240
               Width           =   1455
            End
            Begin VB.Label lblInfo 
               Caption         =   "Right:"
               Height          =   255
               Index           =   7
               Left            =   360
               TabIndex        =   24
               Top             =   960
               Width           =   1215
            End
            Begin VB.Label lblInfo 
               Caption         =   "Left:"
               Height          =   255
               Index           =   6
               Left            =   360
               TabIndex        =   23
               Top             =   720
               Width           =   1215
            End
            Begin VB.Label lblInfo 
               Caption         =   "Down:"
               Height          =   255
               Index           =   5
               Left            =   360
               TabIndex        =   22
               Top             =   480
               Width           =   1215
            End
            Begin VB.Label lblInfo 
               Caption         =   "Up:"
               Height          =   255
               Index           =   4
               Left            =   360
               TabIndex        =   21
               Top             =   240
               Width           =   1215
            End
         End
         Begin VB.OptionButton optInput 
            Caption         =   "Gamepad / Joystick"
            Height          =   255
            Index           =   1
            Left            =   1560
            TabIndex        =   19
            Top             =   480
            Width           =   1935
         End
         Begin VB.OptionButton optInput 
            Caption         =   "Keyboard"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   18
            Top             =   480
            Width           =   1695
         End
      End
      Begin VB.Frame fmeOption 
         Caption         =   "Video"
         Height          =   1695
         Index           =   1
         Left            =   240
         TabIndex        =   10
         Top             =   1920
         Width           =   3735
         Begin VB.CheckBox chkVideo 
            Caption         =   "Use animated backgrounds"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   15
            Top             =   720
            Width           =   3375
         End
         Begin VB.CheckBox chkVideo 
            Caption         =   "Letterbox graphics"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   3375
         End
         Begin VB.Label lblInfo 
            Caption         =   "Available only on standard displays"
            Height          =   255
            Index           =   2
            Left            =   600
            TabIndex        =   14
            Top             =   480
            Width           =   3015
         End
      End
      Begin VB.Frame fmeOption 
         Caption         =   "Audio"
         Height          =   855
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   960
         Width           =   3735
         Begin VB.CheckBox chkAudio 
            Caption         =   "Play sound effects"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   12
            Top             =   480
            Width           =   2055
         End
         Begin VB.CheckBox chkAudio 
            Caption         =   "Play background music"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.CommandButton btnProfile 
         Caption         =   "Delete Selected Profile"
         Height          =   255
         Index           =   1
         Left            =   -72960
         TabIndex        =   6
         Top             =   3480
         Width           =   1815
      End
      Begin VB.ListBox lstProfiles 
         Height          =   1620
         Left            =   -74760
         TabIndex        =   5
         Top             =   1800
         Width           =   3735
      End
      Begin VB.CommandButton btnProfile 
         Caption         =   "Add New Profile"
         Height          =   255
         Index           =   0
         Left            =   -72960
         TabIndex        =   4
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox txtProfileName 
         Height          =   285
         Left            =   -74760
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   1080
         Width           =   3735
      End
      Begin VB.Label lblInfo 
         Caption         =   "Configure the controls for Treasure Quest"
         Height          =   255
         Index           =   3
         Left            =   -74760
         TabIndex        =   16
         Top             =   480
         Width           =   3735
      End
      Begin VB.Label lblInfo 
         Caption         =   "Configure how Treasure Quest looks and sounds"
         Height          =   495
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   3735
      End
      Begin VB.Label lblInfo 
         Caption         =   "Profiles are used to keep track of a player's progress in Treasture Quest.  Add or delete profiles here."
         Height          =   495
         Index           =   0
         Left            =   -74760
         TabIndex        =   7
         Top             =   480
         Width           =   3735
      End
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Top             =   4080
      Width           =   1335
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
