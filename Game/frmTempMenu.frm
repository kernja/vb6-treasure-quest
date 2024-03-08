VERSION 5.00
Begin VB.Form frmTempMenu 
   Caption         =   "Form1"
   ClientHeight    =   3795
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3045
   LinkTopic       =   "Form1"
   ScaleHeight     =   3795
   ScaleWidth      =   3045
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optView 
      Caption         =   "Standard"
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   6
      Top             =   3240
      Value           =   -1  'True
      Width           =   1335
   End
   Begin VB.OptionButton optView 
      Caption         =   "Letterbox"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton btnCommand 
      Caption         =   "Four Player Battle"
      Height          =   495
      Index           =   5
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   2775
   End
   Begin VB.CommandButton btnCommand 
      Caption         =   "Three Player Battle"
      Height          =   495
      Index           =   4
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   2775
   End
   Begin VB.CommandButton btnCommand 
      Caption         =   "Two Player Battle"
      Height          =   495
      Index           =   3
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   2775
   End
   Begin VB.CommandButton btnCommand 
      Caption         =   "Two Player CO-OP"
      Height          =   495
      Index           =   2
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   2775
   End
   Begin VB.CommandButton btnCommand 
      Caption         =   "One Player"
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "frmTempMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCommand_Click(index As Integer)
    If index = 1 Then
        If optView(1).value = True Then
            frmGame.initPlayfield modeOnePlayer, screenStretched
        Else
            frmGame.initPlayfield modeOnePlayer, screenLetterbox
        End If
    ElseIf index = 2 Then
       If optView(1).value = True Then
            frmGame.initPlayfield modetwoplayercoop, screenStretched
        Else
            frmGame.initPlayfield modetwoplayercoop, screenLetterbox
        End If
    ElseIf index = 3 Then
       If optView(1).value = True Then
            frmGame.initPlayfield modeTwoPlayerVS, screenStretched
        Else
            frmGame.initPlayfield modeTwoPlayerVS, screenLetterbox
        End If
    ElseIf index = 4 Then
       If optView(1).value = True Then
            frmGame.initPlayfield modeThrPlayerVS, screenStretched
        Else
            frmGame.initPlayfield modeThrPlayerVS, screenLetterbox
        End If
    ElseIf index = 5 Then
       If optView(1).value = True Then
            frmGame.initPlayfield modeFouPlayerVS, screenStretched
        Else
            frmGame.initPlayfield modeFouPlayerVS, screenLetterbox
        End If
    End If
End Sub

Private Sub Form_Load()
    If mdlMain.isScreenWidescreen = False Then
        optView(0).enabled = False
    Else
        optView(0).enabled = True
    End If
End Sub

