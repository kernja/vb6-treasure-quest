VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.Form frmSprites 
   Caption         =   "Form1"
   ClientHeight    =   3510
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10770
   LinkTopic       =   "Form1"
   ScaleHeight     =   234
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   718
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picContainer 
      Height          =   1215
      Index           =   8
      Left            =   240
      ScaleHeight     =   77
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   77
      TabIndex        =   20
      Top             =   2760
      Width           =   1215
      Begin VB.PictureBox picSimpleBG 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   120
         ScaleHeight     =   615
         ScaleWidth      =   735
         TabIndex        =   21
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.PictureBox picContainer 
      Height          =   1215
      Index           =   14
      Left            =   10200
      ScaleHeight     =   77
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   77
      TabIndex        =   19
      Top             =   120
      Width           =   1215
      Begin VB.PictureBox picPlayerSprite 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1800
         Index           =   7
         Left            =   240
         Picture         =   "frmSprites.frx":0000
         ScaleHeight     =   1800
         ScaleWidth      =   4200
         TabIndex        =   29
         Top             =   600
         Width           =   4200
      End
      Begin VB.PictureBox picPlayerSprite 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1800
         Index           =   6
         Left            =   0
         Picture         =   "frmSprites.frx":18A02
         ScaleHeight     =   1800
         ScaleWidth      =   4200
         TabIndex        =   28
         Top             =   360
         Width           =   4200
      End
   End
   Begin VB.PictureBox picContainer 
      Height          =   1215
      Index           =   12
      Left            =   9480
      ScaleHeight     =   77
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   77
      TabIndex        =   18
      Top             =   120
      Width           =   1215
      Begin VB.PictureBox picPlayerSprite 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1800
         Index           =   5
         Left            =   240
         Picture         =   "frmSprites.frx":31404
         ScaleHeight     =   1800
         ScaleWidth      =   4200
         TabIndex        =   27
         Top             =   720
         Width           =   4200
      End
      Begin VB.PictureBox picPlayerSprite 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1800
         Index           =   4
         Left            =   120
         Picture         =   "frmSprites.frx":49E06
         ScaleHeight     =   1800
         ScaleWidth      =   4200
         TabIndex        =   26
         Top             =   360
         Width           =   4200
      End
   End
   Begin VB.PictureBox picContainer 
      Height          =   1215
      Index           =   10
      Left            =   8760
      ScaleHeight     =   77
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   77
      TabIndex        =   17
      Top             =   120
      Width           =   1215
      Begin VB.PictureBox picPlayerSprite 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1800
         Index           =   3
         Left            =   240
         Picture         =   "frmSprites.frx":62808
         ScaleHeight     =   1800
         ScaleWidth      =   4200
         TabIndex        =   25
         Top             =   600
         Width           =   4200
      End
      Begin VB.PictureBox picPlayerSprite 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1800
         Index           =   2
         Left            =   120
         Picture         =   "frmSprites.frx":7B20A
         ScaleHeight     =   1800
         ScaleWidth      =   4200
         TabIndex        =   24
         Top             =   360
         Width           =   4200
      End
   End
   Begin VB.PictureBox picContainer 
      Height          =   1215
      Index           =   6
      Left            =   7920
      ScaleHeight     =   77
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   77
      TabIndex        =   16
      Top             =   120
      Width           =   1215
      Begin VB.PictureBox picPlayerSprite 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1800
         Index           =   1
         Left            =   120
         Picture         =   "frmSprites.frx":93C0C
         ScaleHeight     =   1800
         ScaleWidth      =   4200
         TabIndex        =   22
         Top             =   480
         Width           =   4200
      End
      Begin VB.PictureBox picPlayerSprite 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1800
         Index           =   0
         Left            =   0
         Picture         =   "frmSprites.frx":AC60E
         ScaleHeight     =   1800
         ScaleWidth      =   4200
         TabIndex        =   23
         Top             =   360
         Width           =   4200
      End
   End
   Begin VB.PictureBox picContainer 
      Height          =   1215
      Index           =   9
      Left            =   5520
      ScaleHeight     =   77
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   13
      Top             =   1440
      Width           =   1095
      Begin VB.PictureBox picInteractiveSprites 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   960
         Index           =   1
         Left            =   360
         Picture         =   "frmSprites.frx":C5010
         ScaleHeight     =   960
         ScaleWidth      =   7680
         TabIndex        =   14
         Top             =   360
         Width           =   7680
      End
      Begin VB.PictureBox picInteractiveSprites 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   960
         Index           =   0
         Left            =   120
         Picture         =   "frmSprites.frx":DD052
         ScaleHeight     =   960
         ScaleWidth      =   7680
         TabIndex        =   15
         Top             =   120
         Width           =   7680
      End
   End
   Begin VB.PictureBox picContainer 
      Height          =   1215
      Index           =   7
      Left            =   5520
      ScaleHeight     =   77
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   77
      TabIndex        =   12
      Top             =   120
      Width           =   1215
      Begin PicClip.PictureClip pcInteractiveSprites 
         Left            =   120
         Top             =   120
         _ExtentX        =   13547
         _ExtentY        =   1693
         _Version        =   393216
         Cols            =   8
         Picture         =   "frmSprites.frx":F5094
      End
   End
   Begin VB.PictureBox picContainer 
      Height          =   1215
      Index           =   5
      Left            =   2880
      ScaleHeight     =   77
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   77
      TabIndex        =   11
      Top             =   120
      Width           =   1215
      Begin PicClip.PictureClip pcEnemySprites 
         Left            =   240
         Top             =   120
         _ExtentX        =   5080
         _ExtentY        =   1693
         _Version        =   393216
         Cols            =   3
         Picture         =   "frmSprites.frx":10D0E6
      End
   End
   Begin VB.PictureBox picContainer 
      Height          =   1215
      Index           =   4
      Left            =   1560
      ScaleHeight     =   77
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   77
      TabIndex        =   10
      Top             =   120
      Width           =   1215
      Begin PicClip.PictureClip pcScenerySprites 
         Left            =   120
         Top             =   120
         _ExtentX        =   13547
         _ExtentY        =   1693
         _Version        =   393216
         Cols            =   8
         Picture         =   "frmSprites.frx":116138
      End
   End
   Begin VB.PictureBox picContainer 
      Height          =   1215
      Index           =   3
      Left            =   240
      ScaleHeight     =   77
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   77
      TabIndex        =   9
      Top             =   120
      Width           =   1215
      Begin PicClip.PictureClip pcPlayfieldSprites 
         Left            =   120
         Top             =   120
         _ExtentX        =   13547
         _ExtentY        =   1693
         _Version        =   393216
         Cols            =   8
         Picture         =   "frmSprites.frx":12E18A
      End
   End
   Begin VB.PictureBox picContainer 
      Height          =   1095
      Index           =   2
      Left            =   2880
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   61
      TabIndex        =   6
      Top             =   1440
      Width           =   975
      Begin VB.PictureBox picEnemySprites 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   7680
         Index           =   1
         Left            =   360
         Picture         =   "frmSprites.frx":1461DC
         ScaleHeight     =   512
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   192
         TabIndex        =   7
         Top             =   480
         Width           =   2880
      End
      Begin VB.PictureBox picEnemySprites 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   7680
         Index           =   0
         Left            =   120
         Picture         =   "frmSprites.frx":18E21E
         ScaleHeight     =   512
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   192
         TabIndex        =   8
         Top             =   240
         Width           =   2880
      End
   End
   Begin VB.PictureBox picContainer 
      Height          =   1215
      Index           =   0
      Left            =   240
      ScaleHeight     =   77
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   3
      Top             =   1440
      Width           =   1095
      Begin VB.PictureBox picPlayfieldSprites 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   960
         Index           =   1
         Left            =   360
         Picture         =   "frmSprites.frx":1D6260
         ScaleHeight     =   960
         ScaleWidth      =   7680
         TabIndex        =   5
         Top             =   360
         Width           =   7680
      End
      Begin VB.PictureBox picPlayfieldSprites 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   960
         Index           =   0
         Left            =   120
         Picture         =   "frmSprites.frx":1EE2A2
         ScaleHeight     =   960
         ScaleWidth      =   7680
         TabIndex        =   4
         Top             =   120
         Width           =   7680
      End
   End
   Begin VB.PictureBox picContainer 
      Height          =   1095
      Index           =   1
      Left            =   1560
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   61
      TabIndex        =   0
      Top             =   1440
      Width           =   975
      Begin VB.PictureBox picScenerySprites 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   960
         Index           =   1
         Left            =   360
         Picture         =   "frmSprites.frx":2062E4
         ScaleHeight     =   64
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   512
         TabIndex        =   2
         Top             =   480
         Width           =   7680
      End
      Begin VB.PictureBox picScenerySprites 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   960
         Index           =   0
         Left            =   120
         Picture         =   "frmSprites.frx":21E326
         ScaleHeight     =   64
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   512
         TabIndex        =   1
         Top             =   240
         Width           =   7680
      End
   End
End
Attribute VB_Name = "frmSprites"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
