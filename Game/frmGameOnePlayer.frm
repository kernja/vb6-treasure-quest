VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form frmGame 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Treasure Quest"
   ClientHeight    =   6285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9390
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   419
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   626
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MCI.MMControl mediaPlayer 
      Height          =   615
      Left            =   480
      TabIndex        =   4
      Top             =   5520
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   1085
      _Version        =   393216
      DeviceType      =   "Sequencer"
      FileName        =   "D:\Jeff\My Documents\Treasure Quest\program\midi\AzureLake.mid"
   End
   Begin VB.PictureBox picRender 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   1335
      Index           =   3
      Left            =   2520
      ScaleHeight     =   89
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   153
      TabIndex        =   3
      Top             =   1680
      Width           =   2295
   End
   Begin VB.PictureBox picRender 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   1335
      Index           =   2
      Left            =   120
      ScaleHeight     =   89
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   153
      TabIndex        =   2
      Top             =   1680
      Width           =   2295
   End
   Begin VB.PictureBox picRender 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   1335
      Index           =   1
      Left            =   2520
      ScaleHeight     =   89
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   153
      TabIndex        =   1
      Top             =   240
      Width           =   2295
   End
   Begin VB.PictureBox picRender 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   1335
      Index           =   0
      Left            =   120
      ScaleHeight     =   89
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   153
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Public Sub initPlayfield(passedGamemode As gameMode, passedScreenType As gameScreenMode)
    Dim i As Single
    If passedScreenType = screenLetterbox Then

        
        If passedGamemode = modeOnePlayer Or passedGamemode = modetwoplayercoop Then
            For i = 1 To 3
                picRender(i).Visible = False
            Next i
        
                picRender(0).height = (Screen.height / Screen.TwipsPerPixelY)
                picRender(0).height = ((picRender(0).height \ 4) * 4)
                
                picRender(0).width = picRender(0).height * 1.333
                picRender(0).width = ((picRender(0).width \ 2) * 2)
                
                picRender(0).Left = ((Screen.width / Screen.TwipsPerPixelX) - picRender(0).width) * 0.5
                picRender(0).Top = 0
                    
        ElseIf passedGamemode = modeTwoPlayerVS Then
            For i = 2 To 3
                picRender(i).Visible = False
            Next i
                
                picRender(0).height = (Screen.height / Screen.TwipsPerPixelY)
                picRender(0).height = ((picRender(0).height \ 4) * 4) * 0.5
                
                picRender(0).width = picRender(0).height * 1.333
                picRender(0).width = ((picRender(0).width \ 2) * 2)
                
                picRender(0).Left = ((Screen.width / Screen.TwipsPerPixelX) - picRender(0).width) * 0.5
                picRender(0).Top = 0
                
                With picRender(1)
                    .Top = picRender(0).Top + picRender(0).height
                    .Left = picRender(0).Left
                    .width = picRender(0).width
                    .height = picRender(0).height
                End With
        ElseIf passedGamemode = modeThrPlayerVS Then
            picRender(3).Visible = False
                
                picRender(0).height = (Screen.height / Screen.TwipsPerPixelY)
                picRender(0).height = ((picRender(0).height \ 4) * 4) * 0.5
                
                picRender(0).width = picRender(0).height * 1.333
                picRender(0).width = ((picRender(0).width \ 2) * 2) * 0.5
                
                picRender(0).Left = ((Screen.width / Screen.TwipsPerPixelX) - picRender(0).width) * 0.5
                picRender(0).Top = 0
                
                
                With picRender(1)
                    .Top = picRender(0).Top
                    .Left = picRender(0).Left + picRender(0).width
                    .width = picRender(0).width
                    .height = picRender(0).height
                End With
                
                With picRender(2)
                    .Top = picRender(0).Top + picRender(0).height
                    .Left = picRender(0).Left
                    .width = picRender(0).width
                    .height = picRender(0).height
                End With
        ElseIf passedGamemode = modeFouPlayerVS Then
            picRender(0).height = (Screen.height / Screen.TwipsPerPixelY)
                picRender(0).height = ((picRender(0).height \ 4) * 4)
                
                picRender(0).width = picRender(0).height * 1.333
                picRender(0).width = ((picRender(0).width \ 2) * 2)
                
                picRender(0).Left = ((Screen.width / Screen.TwipsPerPixelX) - picRender(0).width) * 0.5
                picRender(0).Top = 0
                
                With picRender(1)
                    .Top = picRender(0).Top + picRender(0).height
                    .Left = picRender(0).Left
                    .width = picRender(0).width
                    .height = picRender(0).height
                End With
                
                With picRender(2)
                    .Top = picRender(0).Top + picRender(0).height
                    .Left = picRender(0).Left
                    .width = picRender(0).width
                    .height = picRender(0).height
                End With
                
                With picRender(3)
                    .Top = picRender(0).Top + picRender(0).height
                    .Left = picRender(0).Left + picRender(0).width
                    .width = picRender(0).width
                    .height = picRender(0).height
                End With
        End If
    Else
            If passedGamemode = modeOnePlayer Or passedGamemode = modetwoplayercoop Then
                For i = 1 To 3
                    picRender(i).Visible = False
                Next i
            
                    picRender(0).height = (Screen.height / Screen.TwipsPerPixelY)
                    picRender(0).width = (Screen.width / Screen.TwipsPerPixelX)
                    
                    picRender(0).Left = 0
                    picRender(0).Top = 0
            ElseIf passedGamemode = modeTwoPlayerVS Then
                For i = 2 To 3
                    picRender(i).Visible = False
                Next i
                
                    picRender(0).height = (Screen.height / Screen.TwipsPerPixelY) * 0.5
                    picRender(0).width = (Screen.width / Screen.TwipsPerPixelX)
                    
                    picRender(0).Left = 0
                    picRender(0).Top = 0
                    
                    picRender(1).Left = 0
                    picRender(1).Top = picRender(0).height
                    picRender(1).height = picRender(0).height
                    picRender(1).width = picRender(0).width
            ElseIf passedGamemode = modeThrPlayerVS Then
                picRender(3).Visible = False
                
                   picRender(0).height = (Screen.height / Screen.TwipsPerPixelY) * 0.5
                    picRender(0).width = (Screen.width / Screen.TwipsPerPixelX) * 0.5
                    
                    picRender(0).Left = 0
                    picRender(0).Top = 0
                    
                    picRender(1).Top = 0
                    picRender(1).Left = picRender(0).width
                    picRender(1).height = picRender(0).height
                    picRender(1).width = picRender(0).width
                    
                    picRender(2).Left = 0
                    picRender(2).Top = picRender(0).height
                    picRender(2).height = picRender(0).height
                    picRender(2).width = picRender(0).width
            Else
                    picRender(0).height = (Screen.height / Screen.TwipsPerPixelY) * 0.5
                    picRender(0).width = (Screen.width / Screen.TwipsPerPixelX) * 0.5
                    
                    picRender(0).Left = 0
                    picRender(0).Top = 0
                    
                    picRender(1).Top = 0
                    picRender(1).Left = picRender(0).width
                    picRender(1).height = picRender(0).height
                    picRender(1).width = picRender(0).width
                    
                    picRender(2).Left = 0
                    picRender(2).Top = picRender(0).height
                    picRender(2).height = picRender(0).height
                    picRender(2).width = picRender(0).width
                    
                    picRender(3).Top = picRender(0).height
                    picRender(3).Left = picRender(0).width
                    picRender(3).height = picRender(0).height
                    picRender(3).width = picRender(0).width
            End If
     End If

        
        
        'mainRenderOffset(0) = picRender(0).Width * mainRenderStretch(0) * 0.5
        'mainRenderOffset(1) = picRender(0).Height * mainRenderStretch(1) * 0.5
    
            If passedGamemode = modeOnePlayer Then
                
                
                mainRenderStretch(0) = picRender(0).width / 1024
                mainRenderStretch(1) = picRender(0).height / 768
                
                mdlMain.loadMap , 0
            ElseIf passedGamemode = modeTwoPlayerVS Then
                
                
                mainRenderStretch(0) = picRender(0).width / 1024
                mainRenderStretch(1) = (picRender(0).height / 768) * 2
                
                mdlMain.loadMap , 0
                mdlMain.loadMap , 1
            ElseIf passedGamemode = modeThrPlayerVS Then
             mainRenderStretch(0) = (picRender(0).width / 1024) * 2
                mainRenderStretch(1) = (picRender(0).height / 768) * 2
                
                 mdlMain.loadMap , 0
                mdlMain.loadMap , 1
                mdlMain.loadMap , 2
               
            ElseIf passedGamemode = modeFouPlayerVS Then
                
            
                mainRenderStretch(0) = (picRender(0).width / 1024) * 2
                mainRenderStretch(1) = (picRender(0).height / 768) * 2
                
                 mdlMain.loadMap , 0
                mdlMain.loadMap , 1
                 mdlMain.loadMap , 2
                mdlMain.loadMap , 3
            End If
   
    currentGameMode = passedGamemode
    currentScreenMode = passedScreenType
    frmTempMenu.Hide
    Me.Show
        mdlMain.progRender = True
        mdlMain.renderMe
End Sub


Private Sub Form_Load()
    'mediaPlayer.Command = "Open"
      '  frmSprites.Show
End Sub

Private Sub mediaPlayer_Done(NotifyCode As Integer)
    If loopMusic = True Then
        mediaPlayer.Command = "Play"
    End If
End Sub

Private Sub picRender_Click(index As Integer)
   progExit = True
End Sub
