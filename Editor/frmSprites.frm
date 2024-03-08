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
         Picture         =   "frmSprites.frx":0000
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
         Picture         =   "frmSprites.frx":18042
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
         Picture         =   "frmSprites.frx":30084
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
         Left            =   120
         Top             =   120
         _ExtentX        =   5080
         _ExtentY        =   1693
         _Version        =   393216
         Cols            =   3
         Picture         =   "frmSprites.frx":480D6
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
         Picture         =   "frmSprites.frx":51128
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
         Picture         =   "frmSprites.frx":6917A
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
         Picture         =   "frmSprites.frx":811CC
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
         Picture         =   "frmSprites.frx":C920E
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
         Picture         =   "frmSprites.frx":111250
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
         Picture         =   "frmSprites.frx":129292
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
         Picture         =   "frmSprites.frx":1412D4
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
         Picture         =   "frmSprites.frx":159316
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
