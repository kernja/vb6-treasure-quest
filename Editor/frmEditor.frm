VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmEditor 
   Caption         =   "Form1"
   ClientHeight    =   6285
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   12090
   LinkTopic       =   "Form1"
   ScaleHeight     =   419
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   806
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtMapName 
      Height          =   285
      Left            =   1320
      TabIndex        =   29
      Top             =   120
      Width           =   6495
   End
   Begin VB.Frame frmObjects 
      Caption         =   "Available Objects"
      Enabled         =   0   'False
      Height          =   1455
      Left            =   9360
      TabIndex        =   8
      Top             =   4800
      Width           =   2655
      Begin VB.ListBox lstObjects 
         Height          =   1035
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1335
      End
      Begin VB.Image imagePreview 
         Appearance      =   0  'Flat
         Height          =   855
         Left            =   1560
         Stretch         =   -1  'True
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Preview:"
         Height          =   255
         Index           =   0
         Left            =   1560
         TabIndex        =   12
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame frmViews 
      Caption         =   "Views"
      Height          =   855
      Left            =   7920
      TabIndex        =   2
      Top             =   120
      Width           =   4095
      Begin VB.OptionButton optView 
         Caption         =   "fg2"
         Height          =   255
         Index           =   6
         Left            =   2760
         TabIndex        =   11
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton optView 
         Caption         =   "intrv."
         Height          =   255
         Index           =   2
         Left            =   840
         TabIndex        =   14
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton optView 
         Caption         =   "fg1"
         Height          =   255
         Index           =   5
         Left            =   1800
         TabIndex        =   7
         Top             =   480
         Width           =   615
      End
      Begin VB.OptionButton optView 
         Caption         =   "playfield"
         Height          =   255
         Index           =   4
         Left            =   1800
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optView 
         Caption         =   "pl/enemy"
         Height          =   255
         Index           =   3
         Left            =   840
         TabIndex        =   5
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton optView 
         Caption         =   "bg1"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   735
      End
      Begin VB.OptionButton optView 
         Caption         =   "bg2"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame frmItems 
      Caption         =   "Items"
      Enabled         =   0   'False
      Height          =   5175
      Left            =   7920
      TabIndex        =   1
      Top             =   1080
      Width           =   4095
      Begin TabDlg.SSTab tabFlags 
         Height          =   1455
         Left            =   120
         TabIndex        =   31
         Top             =   1680
         Width           =   3885
         _ExtentX        =   6853
         _ExtentY        =   2566
         _Version        =   393216
         Tabs            =   4
         TabsPerRow      =   4
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Solid"
         TabPicture(0)   =   "frmEditor.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "chkSolid(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "chkSolid(1)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "chkSolid(2)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "chkSolid(3)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "chkSolid(4)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "chkSolid(5)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "chkSolid(6)"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "chkSolid(7)"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "chkSolid(8)"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "chkSolid(9)"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "chkSolid(10)"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "chkSolid(11)"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "chkSolid(12)"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "chkSolid(13)"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "chkSolid(14)"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "chkSolid(15)"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).ControlCount=   16
         TabCaption(1)   =   "Interactive"
         TabPicture(1)   =   "frmEditor.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "chkInteractive(15)"
         Tab(1).Control(1)=   "chkInteractive(14)"
         Tab(1).Control(2)=   "chkInteractive(13)"
         Tab(1).Control(3)=   "chkInteractive(12)"
         Tab(1).Control(4)=   "chkInteractive(11)"
         Tab(1).Control(5)=   "chkInteractive(10)"
         Tab(1).Control(6)=   "chkInteractive(9)"
         Tab(1).Control(7)=   "chkInteractive(8)"
         Tab(1).Control(8)=   "chkInteractive(7)"
         Tab(1).Control(9)=   "chkInteractive(6)"
         Tab(1).Control(10)=   "chkInteractive(5)"
         Tab(1).Control(11)=   "chkInteractive(4)"
         Tab(1).Control(12)=   "chkInteractive(3)"
         Tab(1).Control(13)=   "chkInteractive(2)"
         Tab(1).Control(14)=   "chkInteractive(1)"
         Tab(1).Control(15)=   "chkInteractive(0)"
         Tab(1).ControlCount=   16
         TabCaption(2)   =   "Interactv 2"
         TabPicture(2)   =   "frmEditor.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "chkInteractive(16)"
         Tab(2).Control(1)=   "chkInteractive(17)"
         Tab(2).Control(2)=   "chkInteractive(18)"
         Tab(2).Control(3)=   "chkInteractive(19)"
         Tab(2).Control(4)=   "chkInteractive(20)"
         Tab(2).Control(5)=   "chkInteractive(21)"
         Tab(2).Control(6)=   "chkInteractive(22)"
         Tab(2).Control(7)=   "chkInteractive(23)"
         Tab(2).Control(8)=   "chkInteractive(24)"
         Tab(2).Control(9)=   "chkInteractive(25)"
         Tab(2).Control(10)=   "chkInteractive(26)"
         Tab(2).Control(11)=   "chkInteractive(27)"
         Tab(2).Control(12)=   "chkInteractive(28)"
         Tab(2).Control(13)=   "chkInteractive(29)"
         Tab(2).Control(14)=   "chkInteractive(30)"
         Tab(2).Control(15)=   "chkInteractive(31)"
         Tab(2).ControlCount=   16
         TabCaption(3)   =   "Enemy"
         TabPicture(3)   =   "frmEditor.frx":0054
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "optEnemy(1)"
         Tab(3).Control(1)=   "optEnemy(2)"
         Tab(3).Control(2)=   "optEnemy(3)"
         Tab(3).Control(3)=   "optEnemy(4)"
         Tab(3).Control(4)=   "optEnemy(0)"
         Tab(3).Control(5)=   "optEnemy(5)"
         Tab(3).Control(6)=   "optEnemy(6)"
         Tab(3).Control(7)=   "optEnemy(7)"
         Tab(3).Control(8)=   "optEnemy(8)"
         Tab(3).Control(9)=   "optEnemy(9)"
         Tab(3).Control(10)=   "optEnemy(10)"
         Tab(3).Control(11)=   "optEnemy(11)"
         Tab(3).Control(12)=   "optEnemy(12)"
         Tab(3).Control(13)=   "optEnemy(13)"
         Tab(3).Control(14)=   "optEnemy(14)"
         Tab(3).ControlCount=   15
         Begin VB.OptionButton optEnemy 
            Caption         =   "____"
            Height          =   195
            Index           =   14
            Left            =   -72120
            TabIndex        =   94
            Top             =   1080
            Width           =   855
         End
         Begin VB.OptionButton optEnemy 
            Caption         =   "____"
            Height          =   195
            Index           =   13
            Left            =   -72120
            TabIndex        =   93
            Top             =   840
            Width           =   855
         End
         Begin VB.OptionButton optEnemy 
            Caption         =   "____"
            Height          =   195
            Index           =   12
            Left            =   -72120
            TabIndex        =   92
            Top             =   600
            Width           =   855
         End
         Begin VB.OptionButton optEnemy 
            Caption         =   "____"
            Height          =   195
            Index           =   11
            Left            =   -72120
            TabIndex        =   91
            Top             =   360
            Width           =   855
         End
         Begin VB.OptionButton optEnemy 
            Caption         =   "____"
            Height          =   195
            Index           =   10
            Left            =   -72960
            TabIndex        =   90
            Top             =   1080
            Width           =   855
         End
         Begin VB.OptionButton optEnemy 
            Caption         =   "____"
            Height          =   195
            Index           =   9
            Left            =   -72960
            TabIndex        =   89
            Top             =   840
            Width           =   855
         End
         Begin VB.OptionButton optEnemy 
            Caption         =   "____"
            Height          =   195
            Index           =   8
            Left            =   -72960
            TabIndex        =   88
            Top             =   600
            Width           =   855
         End
         Begin VB.OptionButton optEnemy 
            Caption         =   "____"
            Height          =   195
            Index           =   7
            Left            =   -72960
            TabIndex        =   87
            Top             =   360
            Width           =   855
         End
         Begin VB.OptionButton optEnemy 
            Caption         =   "____"
            Height          =   195
            Index           =   6
            Left            =   -73920
            TabIndex        =   86
            Top             =   1080
            Width           =   855
         End
         Begin VB.OptionButton optEnemy 
            Caption         =   "____"
            Height          =   195
            Index           =   5
            Left            =   -73920
            TabIndex        =   85
            Top             =   840
            Width           =   855
         End
         Begin VB.CheckBox chkInteractive 
            Caption         =   "___"
            Height          =   255
            Index           =   31
            Left            =   -72120
            TabIndex        =   84
            Top             =   1080
            Width           =   855
         End
         Begin VB.CheckBox chkInteractive 
            Caption         =   "___"
            Height          =   255
            Index           =   30
            Left            =   -72120
            TabIndex        =   83
            Top             =   840
            Width           =   855
         End
         Begin VB.CheckBox chkInteractive 
            Caption         =   "___"
            Height          =   255
            Index           =   29
            Left            =   -72120
            TabIndex        =   82
            Top             =   600
            Width           =   855
         End
         Begin VB.CheckBox chkInteractive 
            Caption         =   "___"
            Height          =   255
            Index           =   28
            Left            =   -72120
            TabIndex        =   81
            Top             =   360
            Width           =   855
         End
         Begin VB.CheckBox chkInteractive 
            Caption         =   "___"
            Height          =   255
            Index           =   27
            Left            =   -73080
            TabIndex        =   80
            Top             =   1080
            Width           =   855
         End
         Begin VB.CheckBox chkInteractive 
            Caption         =   "___"
            Height          =   255
            Index           =   26
            Left            =   -73080
            TabIndex        =   79
            Top             =   840
            Width           =   855
         End
         Begin VB.CheckBox chkInteractive 
            Caption         =   "___"
            Height          =   255
            Index           =   25
            Left            =   -73080
            TabIndex        =   78
            Top             =   600
            Width           =   855
         End
         Begin VB.CheckBox chkInteractive 
            Caption         =   "___"
            Height          =   255
            Index           =   24
            Left            =   -73080
            TabIndex        =   77
            Top             =   360
            Width           =   855
         End
         Begin VB.CheckBox chkInteractive 
            Caption         =   "___"
            Height          =   255
            Index           =   23
            Left            =   -73920
            TabIndex        =   76
            Top             =   1080
            Width           =   855
         End
         Begin VB.CheckBox chkInteractive 
            Caption         =   "___"
            Height          =   255
            Index           =   22
            Left            =   -73920
            TabIndex        =   75
            Top             =   840
            Width           =   855
         End
         Begin VB.CheckBox chkInteractive 
            Caption         =   "___"
            Height          =   255
            Index           =   21
            Left            =   -73920
            TabIndex        =   74
            Top             =   600
            Width           =   855
         End
         Begin VB.CheckBox chkInteractive 
            Caption         =   "___"
            Height          =   255
            Index           =   20
            Left            =   -73920
            TabIndex        =   73
            Top             =   360
            Width           =   855
         End
         Begin VB.CheckBox chkSolid 
            Caption         =   "____"
            Height          =   255
            Index           =   15
            Left            =   3000
            TabIndex        =   72
            Top             =   1080
            Width           =   735
         End
         Begin VB.CheckBox chkSolid 
            Caption         =   "____"
            Height          =   255
            Index           =   14
            Left            =   3000
            TabIndex        =   71
            Top             =   840
            Width           =   735
         End
         Begin VB.CheckBox chkSolid 
            Caption         =   "____"
            Height          =   255
            Index           =   13
            Left            =   3000
            TabIndex        =   70
            Top             =   600
            Width           =   735
         End
         Begin VB.CheckBox chkSolid 
            Caption         =   "____"
            Height          =   255
            Index           =   12
            Left            =   3000
            TabIndex        =   69
            Top             =   360
            Width           =   735
         End
         Begin VB.CheckBox chkSolid 
            Caption         =   "____"
            Height          =   255
            Index           =   11
            Left            =   2280
            TabIndex        =   68
            Top             =   1080
            Width           =   735
         End
         Begin VB.CheckBox chkSolid 
            Caption         =   "____"
            Height          =   255
            Index           =   10
            Left            =   2280
            TabIndex        =   67
            Top             =   840
            Width           =   735
         End
         Begin VB.CheckBox chkSolid 
            Caption         =   "____"
            Height          =   255
            Index           =   9
            Left            =   2280
            TabIndex        =   66
            Top             =   600
            Width           =   735
         End
         Begin VB.OptionButton optEnemy 
            Caption         =   "not an enemy"
            Height          =   435
            Index           =   0
            Left            =   -74880
            TabIndex        =   65
            Top             =   360
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton optEnemy 
            Caption         =   "trckr"
            Height          =   195
            Index           =   4
            Left            =   -73920
            TabIndex        =   64
            Top             =   600
            Width           =   855
         End
         Begin VB.OptionButton optEnemy 
            Caption         =   "wndr f"
            Height          =   195
            Index           =   3
            Left            =   -73920
            TabIndex        =   63
            Top             =   360
            Width           =   855
         End
         Begin VB.OptionButton optEnemy 
            Caption         =   "wndr m"
            Height          =   195
            Index           =   2
            Left            =   -74880
            TabIndex        =   62
            Top             =   1080
            Width           =   855
         End
         Begin VB.OptionButton optEnemy 
            Caption         =   "wndr s"
            Height          =   195
            Index           =   1
            Left            =   -74880
            TabIndex        =   61
            Top             =   840
            Width           =   855
         End
         Begin VB.CheckBox chkInteractive 
            Caption         =   "lvl exit"
            Height          =   255
            Index           =   19
            Left            =   -74880
            TabIndex        =   60
            Top             =   1080
            Width           =   855
         End
         Begin VB.CheckBox chkInteractive 
            Caption         =   "p2 strt"
            Height          =   255
            Index           =   18
            Left            =   -74880
            TabIndex        =   59
            Top             =   840
            Width           =   855
         End
         Begin VB.CheckBox chkInteractive 
            Caption         =   "p1 strt"
            Height          =   255
            Index           =   17
            Left            =   -74880
            TabIndex        =   58
            Top             =   600
            Width           =   855
         End
         Begin VB.CheckBox chkInteractive 
            Caption         =   "1-up"
            Height          =   255
            Index           =   16
            Left            =   -74880
            TabIndex        =   57
            Top             =   360
            Width           =   855
         End
         Begin VB.CheckBox chkInteractive 
            Caption         =   "gem d"
            Height          =   255
            Index           =   15
            Left            =   -72240
            TabIndex        =   56
            Top             =   1080
            Width           =   855
         End
         Begin VB.CheckBox chkInteractive 
            Caption         =   "gem c"
            Height          =   255
            Index           =   14
            Left            =   -72240
            TabIndex        =   55
            Top             =   840
            Width           =   855
         End
         Begin VB.CheckBox chkInteractive 
            Caption         =   "gem b"
            Height          =   255
            Index           =   13
            Left            =   -72240
            TabIndex        =   54
            Top             =   600
            Width           =   855
         End
         Begin VB.CheckBox chkInteractive 
            Caption         =   "gem a"
            Height          =   255
            Index           =   12
            Left            =   -72240
            TabIndex        =   53
            Top             =   360
            Width           =   855
         End
         Begin VB.CheckBox chkInteractive 
            Caption         =   "rq. gem"
            Height          =   255
            Index           =   11
            Left            =   -73200
            TabIndex        =   52
            Top             =   1080
            Width           =   855
         End
         Begin VB.CheckBox chkInteractive 
            Caption         =   "fill l"
            Height          =   255
            Index           =   10
            Left            =   -73200
            TabIndex        =   51
            Top             =   840
            Width           =   855
         End
         Begin VB.CheckBox chkInteractive 
            Caption         =   "fill m"
            Height          =   255
            Index           =   9
            Left            =   -73200
            TabIndex        =   50
            Top             =   600
            Width           =   855
         End
         Begin VB.CheckBox chkInteractive 
            Caption         =   "fill s"
            Height          =   255
            Index           =   8
            Left            =   -73200
            TabIndex        =   49
            Top             =   360
            Width           =   855
         End
         Begin VB.CheckBox chkInteractive 
            Caption         =   "deton."
            Height          =   255
            Index           =   7
            Left            =   -74040
            TabIndex        =   48
            Top             =   1080
            Width           =   855
         End
         Begin VB.CheckBox chkInteractive 
            Caption         =   "1 gas"
            Height          =   255
            Index           =   6
            Left            =   -74040
            TabIndex        =   47
            Top             =   840
            Width           =   855
         End
         Begin VB.CheckBox chkInteractive 
            Caption         =   ".75 gas"
            Height          =   255
            Index           =   5
            Left            =   -74040
            TabIndex        =   46
            Top             =   600
            Width           =   855
         End
         Begin VB.CheckBox chkInteractive 
            Caption         =   ".5 gas"
            Height          =   255
            Index           =   4
            Left            =   -74040
            TabIndex        =   45
            Top             =   360
            Width           =   855
         End
         Begin VB.CheckBox chkInteractive 
            Caption         =   ".25 gas"
            Height          =   255
            Index           =   3
            Left            =   -74880
            TabIndex        =   44
            Top             =   1080
            Width           =   855
         End
         Begin VB.CheckBox chkInteractive 
            Caption         =   "exit"
            Height          =   255
            Index           =   2
            Left            =   -74880
            TabIndex        =   43
            Top             =   840
            Width           =   735
         End
         Begin VB.CheckBox chkInteractive 
            Caption         =   "key"
            Height          =   255
            Index           =   1
            Left            =   -74880
            TabIndex        =   42
            Top             =   600
            Width           =   735
         End
         Begin VB.CheckBox chkInteractive 
            Caption         =   "ladder"
            Height          =   255
            Index           =   0
            Left            =   -74880
            TabIndex        =   41
            Top             =   360
            Width           =   735
         End
         Begin VB.CheckBox chkSolid 
            Caption         =   "lock"
            Height          =   255
            Index           =   8
            Left            =   2280
            TabIndex        =   40
            Top             =   360
            Width           =   735
         End
         Begin VB.CheckBox chkSolid 
            Caption         =   "conv. r/f"
            Height          =   255
            Index           =   7
            Left            =   1200
            TabIndex        =   39
            Top             =   1080
            Width           =   975
         End
         Begin VB.CheckBox chkSolid 
            Caption         =   "conv. l/f"
            Height          =   255
            Index           =   6
            Left            =   1200
            TabIndex        =   38
            Top             =   840
            Width           =   975
         End
         Begin VB.CheckBox chkSolid 
            Caption         =   "conv. r/m"
            Height          =   255
            Index           =   5
            Left            =   1200
            TabIndex        =   37
            Top             =   600
            Width           =   1095
         End
         Begin VB.CheckBox chkSolid 
            Caption         =   "conv. l/m"
            Height          =   255
            Index           =   4
            Left            =   1200
            TabIndex        =   36
            Top             =   360
            Width           =   975
         End
         Begin VB.CheckBox chkSolid 
            Caption         =   "conv. r/s"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   35
            Top             =   1080
            Width           =   975
         End
         Begin VB.CheckBox chkSolid 
            Caption         =   "conv. l/s"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   34
            Top             =   840
            Width           =   975
         End
         Begin VB.CheckBox chkSolid 
            Caption         =   "Hidden"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   33
            Top             =   600
            Width           =   975
         End
         Begin VB.CheckBox chkSolid 
            Caption         =   "NonSolid"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   32
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame frmScale 
         Caption         =   "Scale X"
         Height          =   735
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Top             =   3600
         Width           =   1215
         Begin VB.OptionButton optScaleX 
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   255
         End
         Begin VB.OptionButton optScaleX 
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   24
            Top             =   240
            Value           =   -1  'True
            Width           =   255
         End
         Begin VB.OptionButton optScaleX 
            Height          =   255
            Index           =   2
            Left            =   600
            TabIndex        =   23
            Top             =   240
            Width           =   255
         End
         Begin VB.OptionButton optScaleX 
            Height          =   255
            Index           =   3
            Left            =   840
            TabIndex        =   22
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lblInfo 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   ".5x 1x  2x  4x"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   26
            Top             =   480
            Width           =   1215
         End
      End
      Begin VB.Frame frmScale 
         Caption         =   "Scale Y"
         Height          =   735
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   4320
         Width           =   1215
         Begin VB.OptionButton optScaleY 
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   255
         End
         Begin VB.OptionButton optScaleY 
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   18
            Top             =   240
            Value           =   -1  'True
            Width           =   255
         End
         Begin VB.OptionButton optScaleY 
            Height          =   255
            Index           =   2
            Left            =   600
            TabIndex        =   17
            Top             =   240
            Width           =   255
         End
         Begin VB.OptionButton optScaleY 
            Height          =   255
            Index           =   3
            Left            =   840
            TabIndex        =   16
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lblInfo 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   ".5x 1x  2x  4x"
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   20
            Top             =   480
            Width           =   1215
         End
      End
      Begin VB.ListBox lstItems 
         Height          =   1035
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   3855
      End
      Begin VB.CommandButton btnDelete 
         Caption         =   "Delete Item"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1320
         Width           =   3855
      End
      Begin MSComctlLib.Slider slideSpecID 
         Height          =   255
         Left            =   960
         TabIndex        =   27
         Top             =   3240
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   450
         _Version        =   393216
         Max             =   9
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "spec #"
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   28
         Top             =   3240
         Width           =   1095
      End
   End
   Begin VB.PictureBox picRender 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   5520
      Left            =   120
      ScaleHeight     =   368
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   512
      TabIndex        =   0
      Top             =   480
      Width           =   7680
      Begin VB.Shape shapeHighlight 
         BorderStyle     =   5  'Dash-Dot-Dot
         Height          =   255
         Left            =   120
         Top             =   3000
         Width           =   255
      End
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Map Name"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   30
      Top             =   120
      Width           =   1095
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileMap 
         Caption         =   "New Map"
         Index           =   0
      End
      Begin VB.Menu mnuFileMap 
         Caption         =   "Open Map"
         Index           =   1
      End
      Begin VB.Menu mnuFileMap 
         Caption         =   "Save Map"
         Index           =   2
      End
      Begin VB.Menu mnuFileMap 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuBackground 
      Caption         =   "&Background"
      Begin VB.Menu mnuBackgroundOption 
         Caption         =   "Lake Overview"
         Checked         =   -1  'True
         Index           =   0
      End
   End
   Begin VB.Menu mnuMusic 
      Caption         =   "&Music"
      Begin VB.Menu mnuMusicOption 
         Caption         =   "Azure Lake"
         Checked         =   -1  'True
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub btnDelete_Click()
Dim i As Single
Dim tempItem As classObject
    If lstItems.ListIndex <> -1 Then
        Set tempItem = map.getItem(lstItems.ItemData(lstItems.ListIndex))
            tempItem.destroyObject
        map.setItem lstItems.ItemData(lstItems.ListIndex), tempItem
    End If
    
    refreshObjectListBox
End Sub

Private Sub chkInteractive_Click(Index As Integer)
If lstItems.ListIndex <> -1 Then
    
    Dim tempIndex As Integer
    Dim tempObject As classObject
        tempIndex = lstItems.ItemData(lstItems.ListIndex)
        
        Set tempObject = map.getItem(tempIndex)
        
            If chkInteractive(Index).value = Checked Then
                tempObject.setFlagInteractiveValue Val(Index), True
            Else
                tempObject.setFlagInteractiveValue Val(Index), False
            End If
        map.setItem tempIndex, tempObject
    End If

End Sub

Private Sub chkSolid_Click(Index As Integer)
If lstItems.ListIndex <> -1 Then
    
    Dim tempIndex As Integer
    Dim tempObject As classObject
        tempIndex = lstItems.ItemData(lstItems.ListIndex)
        
        Set tempObject = map.getItem(tempIndex)
        
            If chkSolid(Index).value = Checked Then
                tempObject.setFlagSolidValue Val(Index), True
            Else
                tempObject.setFlagSolidValue Val(Index), False
            End If
        map.setItem tempIndex, tempObject
    End If
End Sub

Private Sub Form_Load()
    Load frmSprites
    'frmSprites.Show
    newMap
End Sub

Private Sub Form_Unload(Cancel As Integer)
    progExit = True
End Sub

Private Sub lstItems_Click()
    If lstItems.ListIndex <> -1 Then
        Dim tempScaleX As Single
        Dim tempScaleY As Single
            tempScaleX = map.getItem(lstItems.ItemData(lstItems.ListIndex)).getScaleX
            tempScaleY = map.getItem(lstItems.ItemData(lstItems.ListIndex)).getScaleY
            
            If tempScaleX = 0.5 Then
                Me.optScaleX(0).value = True
            ElseIf tempScaleX = 1 Then
                Me.optScaleX(1).value = True
            ElseIf tempScaleX = 2 Then
                Me.optScaleX(2).value = True
            Else
                Me.optScaleX(3).value = True
            End If
            
            If tempScaleY = 0.5 Then
                Me.optScaleY(0).value = True
            ElseIf tempScaleY = 1 Then
                Me.optScaleY(1).value = True
            ElseIf tempScaleY = 2 Then
                Me.optScaleY(2).value = True
            Else
                Me.optScaleY(3).value = True
            End If
            
        Dim i As Single
            For i = 0 To map.getItem(lstItems.ItemData(lstItems.ListIndex)).getFlagSolidCount
                If map.getItem(lstItems.ItemData(lstItems.ListIndex)).getFlagSolidValue(i) = True Then
                    Me.chkSolid(i).value = Checked
                Else
                    Me.chkSolid(i).value = Unchecked
                End If
            Next i
            
            For i = 0 To map.getItem(lstItems.ItemData(lstItems.ListIndex)).getFlagInteractiveCount
                If map.getItem(lstItems.ItemData(lstItems.ListIndex)).getFlagInteractiveValue(i) = True Then
                    Me.chkInteractive(i).value = Checked
                Else
                    Me.chkInteractive(i).value = Unchecked
                End If
            Next i
            
            optEnemy((map.getItem(lstItems.ItemData(lstItems.ListIndex)).getFlagEnemy) + 1).value = True
            lstObjects.ListIndex = -1
            imagePreview.Picture = Nothing
    End If
End Sub

Private Sub lstObjects_Click()
    lstItems.ListIndex = -1
    shapeHighlight.Width = 16
    shapeHighlight.Height = 16
    optScaleX(1).value = True
    optScaleY(1).value = True
    
    If lstObjects.ListIndex <> -1 Then
        If optView(2).value = True Then
            imagePreview.Picture = frmSprites.pcInteractiveSprites.GraphicCell(lstObjects.ListIndex)
        ElseIf optView(3).value = True Then
            imagePreview.Picture = frmSprites.pcEnemySprites.GraphicCell(lstObjects.ListIndex)
        ElseIf optView(4).value = True Then
            imagePreview.Picture = frmSprites.pcPlayfieldSprites.GraphicCell(lstObjects.ListIndex)
         ElseIf optView(0).value = True Or optView(1).value = True Or optView(5).value = True Or optView(6).value = True Then
            imagePreview.Picture = frmSprites.pcScenerySprites.GraphicCell(lstObjects.ListIndex)
        End If
    End If
End Sub

Private Sub mnuExit_Click()
    progExit = True
End Sub

Private Sub mnuFileMap_Click(Index As Integer)
    If Index = 0 Then
        mdlMain.newMap
            txtMapName.Text = ""
            refreshObjectListBox
    ElseIf Index = 1 Then
        mdlMain.loadMap
    ElseIf Index = 2 Then
       mdlMain.saveMap
    End If
End Sub

Private Sub optEnemy_Click(Index As Integer)
If lstItems.ListIndex <> -1 Then
    
    Dim tempIndex As Integer
    Dim tempObject As classObject
        tempIndex = lstItems.ItemData(lstItems.ListIndex)
        
        Set tempObject = map.getItem(tempIndex)
        
            tempObject.setFlagEnemy (Index - 1)
        map.setItem tempIndex, tempObject
    End If

End Sub

Private Sub optScaleX_Click(Index As Integer)
If lstItems.ListIndex <> -1 Then
    
    Dim tempIndex As Integer
    Dim tempObject As classObject
        tempIndex = lstItems.ItemData(lstItems.ListIndex)
        
        Set tempObject = map.getItem(tempIndex)
        
        If Index = 0 Then
            tempObject.setScaleX (0.5)
        ElseIf Index = 1 Then
           tempObject.setScaleX (1)
        ElseIf Index = 2 Then
             tempObject.setScaleX (2)
        Else
          tempObject.setScaleX (4)
        End If
        
        map.setItem tempIndex, tempObject
    End If
End Sub

Private Sub optScaleY_Click(Index As Integer)
 If lstItems.ListIndex <> -1 Then
    
    Dim tempIndex As Integer
    Dim tempObject As classObject
        tempIndex = lstItems.ItemData(lstItems.ListIndex)
        
        Set tempObject = map.getItem(tempIndex)
        
        If Index = 0 Then
            tempObject.setScaleY (0.5)
        ElseIf Index = 1 Then
           tempObject.setScaleY (1)
        ElseIf Index = 2 Then
             tempObject.setScaleY (2)
        Else
          tempObject.setScaleY (4)
        End If
        
        map.setItem tempIndex, tempObject
    End If
End Sub

Private Sub optView_Click(Index As Integer)
Dim i As Integer

lstObjects.Clear
imagePreview.Picture = Nothing

refreshObjectListBox
currentSelectionMode = Index
        If Index = 2 Then
            'interactive
            
                For i = 0 To (frmSprites.picInteractiveSprites(0).Width / 64) - 1
                    lstObjects.AddItem ("interactive " & i)
                Next i
            frmObjects.enabled = True
            frmItems.enabled = True
        ElseIf Index = 3 Then
            'enemies
            
                For i = 0 To (frmSprites.picEnemySprites(0).Width / 64) - 1
                    lstObjects.AddItem ("enemies " & i)
                Next i
            frmObjects.enabled = True
            frmItems.enabled = True
        ElseIf Index = 4 Then
            'playfield objects
            
                For i = 0 To (frmSprites.picPlayfieldSprites(0).Width / 64) - 1
                    lstObjects.AddItem ("playfield " & i)
                Next i
            frmObjects.enabled = True
            frmItems.enabled = True
        ElseIf Index <= 1 Or Index >= 5 Then
                    'scenery objects
            
                For i = 0 To (frmSprites.picScenerySprites(0).Width / 64) - 1
                    lstObjects.AddItem ("scenery " & i)
                Next i
            frmObjects.enabled = True
            frmItems.enabled = True
        End If
        
    refreshObjectListBox
End Sub

Private Sub picRender_Click()
Dim selectedLayer As Integer
    selectedLayer = -1
    
    If lstObjects.ListIndex > -1 Then
    If optView(0).value = True Then selectedLayer = gameObjectLayer.layerBg2
    If optView(1).value = True Then selectedLayer = gameObjectLayer.layerbg1
    If optView(2).value = True Then selectedLayer = gameObjectLayer.layerIntv
    If optView(3).value = True Then selectedLayer = gameObjectLayer.layerPe
    If optView(4).value = True Then selectedLayer = gameObjectLayer.layerPf
    If optView(5).value = True Then selectedLayer = gameObjectLayer.layerfg1
    If optView(6).value = True Then selectedLayer = gameObjectLayer.layerfg2
        If selectedLayer = gameObjectLayer.layerBg2 Or selectedLayer = gameObjectLayer.layerbg1 Or selectedLayer = gameObjectLayer.layerfg1 Or selectedLayer = gameObjectLayer.layerfg2 Then
           map.addNewScenery lstObjects.ListIndex, selectedLayer, shapeHighlight.Left * 2, shapeHighlight.Top * 2
        ElseIf selectedLayer = gameObjectLayer.layerIntv Then
            map.addNewInteractive lstObjects.ListIndex, selectedLayer, shapeHighlight.Left * 2, shapeHighlight.Top * 2
        ElseIf selectedLayer = gameObjectLayer.layerPe Then
            map.addNewEnemy lstObjects.ListIndex, selectedLayer, shapeHighlight.Left * 2, shapeHighlight.Top * 2
        ElseIf selectedLayer = gameObjectLayer.layerPf Then
            map.addNewPlayfield lstObjects.ListIndex, selectedLayer, shapeHighlight.Left * 2, shapeHighlight.Top * 2
        End If
        
        
    End If
    
    lstObjects.ListIndex = -1
    refreshObjectListBox
    changeCheckMarks
End Sub

Private Sub picRender_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim tempResult(0 To 1) As Integer
        Dim tempXScale As Single
        Dim tempYScale As Single
        
            If optScaleX(0).value = True Then
                tempXScale = 0.5
            ElseIf optScaleX(1).value = True Then
                tempXScale = 1
            ElseIf optScaleX(2).value = True Then
                tempXScale = 2
            Else
                tempXScale = 3
            End If
            
            
            
            If optScaleY(0).value = True Then
                tempYScale = 0.5
            ElseIf optScaleY(1).value = True Then
                tempYScale = 1
            ElseIf optScaleY(2).value = True Then
                tempYScale = 2
            Else
                tempYScale = 3
            End If
            
            shapeHighlight.Width = 16 * tempXScale
            shapeHighlight.Height = 16 * tempYScale
    tempResult(0) = X / (16 * tempXScale)
    tempResult(1) = Y / (16 * tempYScale)
    
    shapeHighlight.Left = tempResult(0) * 16 * tempXScale
    shapeHighlight.Top = tempResult(1) * 16 * tempYScale
End Sub

Sub refreshObjectListBox()
    Dim i As Integer
    Dim j As Integer
        j = -1
    
    lstItems.Clear
        For i = 0 To map.count
            If map.getItem(i).enabled = True And map.getItem(i).deleteMe = False Then
                If map.getItem(i).objectLayer = currentSelectionMode Then
                    j = j + 1
                    lstItems.AddItem "map item " & map.getItem(i).id, j
                        lstItems.ItemData(j) = i
                End If
            End If
        Next i
        
    
End Sub

Sub changeCheckMarks()
    On Error Resume Next
    
        For i = 0 To 15
            chkSolid(i).value = Unchecked
            'chkInteractive(i).value = Unchecked
        Next i
        
        For i = 0 To 31
            chkSolid(i).value = Unchecked
            chkInteractive(i).value = Unchecked
        Next i
        optEnemy(0).value = True

End Sub

