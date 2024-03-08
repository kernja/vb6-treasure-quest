Attribute VB_Name = "mdlMain"
Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public progRender As Boolean
Public progExit As Boolean
Public currentSelectionMode As Integer

Public currentGameMode As gameMode
Public currentScreenMode As gameScreenMode

Public mainRenderStretch(0 To 1) As Single
Public mainRenderOffset(0 To 1) As Single

Public objFSO As FileSystemObject
Dim ts As TextStream

Public map(0 To 3) As classMap
Public players(0 To 3) As classPlayer
Public currentLoadIndex As Single
Public currentRenderIndex As Single

Public loadedLevelMusic As Boolean
Public loadedLevelSimpleBG As Boolean
Public loopMusic As Boolean

Public Enum gameScreenMode
    screenStretched = 0
    screenLetterbox = 1
End Enum

Public Enum gameMode
    modeOnePlayer = 0
    modetwoplayercoop = 1
    modeTwoPlayerVS = 2
    modeThrPlayerVS = 3
    modeFouPlayerVS = 4
End Enum
Public Type gameObjectWorldData
    posX As Single
    posY As Single
    scaleX As Single
    scaleY As Single
    sizeWidth As Single
    sizeHeight As Single
End Type

Public Enum gameEnemyType
    typeNothing = -1
    typeWandererSlow = 0
    typeWandererMed = 1
    typeWandererFast = 2
    typeTracker = 3
End Enum

Public Enum gameInteractiveType
    typeNothing = -1
End Enum

Public Enum gameObjectType
    typeNothing = -1
    typeScenery = 0
    typePlayfield = 2
    typePlayerEnemy = 1
    typeInteractive = 3
    typeEnemy = 4
    typePlayer = 5
End Enum

Public Enum gameObjectLayer
    layerNothing = -1
    layerBg2 = 0
    layerbg1 = 1
    layerIntv = 2
    layerPe = 3
    layerPf = 4
    layerfg1 = 5
    layerfg2 = 6
End Enum

Sub Main()
    progRender = False
    progExit = False
      Set objFSO = New FileSystemObject
      
        For i = 0 To 3
            Set map(i) = New classMap
        Next i
            
            defineDefaultKeys
            loadedLevelMusic = False
            loadedLevelSimpleBG = False
            loopMusic = True
            frmTempMenu.Show
End Sub

Sub quitProgram()
    Dim i As Single
    
       Set objFSO = Nothing
           
        For i = 0 To 3
            Set map(i) = Nothing
        Next i
        
       End
End Sub

Sub renderMe()
Dim i As Single
On Error Resume Next

    TimeElapse = 0.015
            Do While (progExit = False)
            
                DoEvents
                TimeElapse = TimeElapse * 0.015
                FrameTiming True



            
                If currentGameMode = modeOnePlayer Or currentGameMode = modetwoplayercoop Then
                
                        If progRender = True Then
                            For i = 0 To 0
                                currentRenderIndex = i
                                    frmGame.picRender(i).Cls
                                    map(i).renderMe frmGame.picRender(i).hdc
                            Next i
                        End If
                ElseIf currentGameMode = modeTwoPlayerVS Then
                    
                        If progRender = True Then
                            For i = 0 To 1
                                currentRenderIndex = i
                                frmGame.picRender(i).Cls
                                    map(i).renderMe frmGame.picRender(i).hdc
                            Next i
                        End If

                ElseIf currentGameMode = modeThrPlayerVS Then
                
                        If progRender = True Then
                                 For i = 0 To 2
                                currentRenderIndex = i
                                frmGame.picRender(i).Cls
                                    map(i).renderMe frmGame.picRender(i).hdc
                            Next i
                        End If
                
                ElseIf currentGameMode = modeFouPlayerVS Then
                        If progRender = True Then
                             For i = 0 To 3
                                currentRenderIndex = i
                                frmGame.picRender(i).Cls
                                    map(i).renderMe frmGame.picRender(i).hdc
                            Next i
                        End If
                End If
            
           FrameTiming False
            Loop
    

    quitProgram
End Sub
Sub loadMap(Optional path As String = "map.txt", Optional mapIndex As Single = 0)
    Dim i As Single, j As Single, k As Single, l As Single
    Dim count As Single
    Dim tempItem As classObject
    Dim tempType As Single
    Dim tempScale() As String
    Dim tempInteractiveFlags() As String
    Dim tempPos() As String
    Dim tempSolidFlags() As String
    Dim tempEnemy As Single
    
    Set map(mapIndex) = New classMap
    currentLoadIndex = mapIndex
    Set ts = objFSO.OpenTextFile(path, ForReading, False, TristateFalse)
    
    map(mapIndex).renderOffsetXMax = 640 * mainRenderStretch(0)
    map(mapIndex).renderOffsetX = map(mapIndex).renderOffsetXMax
     map(mapIndex).renderOffsetYMax = 512 * mainRenderStretch(1)
    map(mapIndex).renderOffsetY = map(mapIndex).renderOffsetYMax
    
    'skip the six header lines
    For i = 0 To 5
        ts.SkipLine
    Next i
    'input the map name
        map(mapIndex).mapName = ts.ReadLine
            'frmEditor.txtMapName = map(mapIndex).mapName
    'skip 5 more lines
    For i = 0 To 4
        ts.SkipLine
    Next i
    'input map background
        map(mapIndex).mapBackground = Val(ts.ReadLine)
            'frmEditor.mnuBackgroundOption(map(mapIndex).mapBackground).Checked = True
                loadSimpleBackground (map(mapIndex).mapBackground)
    'skip 4 more lines
    For i = 0 To 4
        ts.SkipLine
    Next i
   'input map music
        map(mapIndex).mapMusic = Val(ts.ReadLine)
            loadGameMusic (map(mapIndex).mapMusic)
            'frmEditor.mnuMusicOption(map(mapIndex).mapBackground).Checked = True
  
    'skip three more lines
        For i = 0 To 2
            ts.SkipLine
        Next i
    For k = 0 To 6
        'skip four more lines
        For i = 0 To 2
            ts.SkipLine 'ts.WriteLine ("//-----------")
        Next i
           count = Val(ts.ReadLine)
                If count > -1 Then
                    For i = 0 To count
                        Set tempItem = New classObject
                        
                        'If tempItem.objectLayer = k Then
                           If i = 0 Then
                            For l = 0 To 4
                             '   MsgBox ts.ReadLine
                             ts.SkipLine
                            Next l
                        Else
                            For l = 0 To 3
                            ts.SkipLine
                            Next l
                        End If
                        
                              '  i 'f k = gameObjectLayer.layerBg2 Then
                ' tempItem.newScenery , k
                                    tempPos = Split(ts.ReadLine, ",")
                                        'MsgBox tempPos
                                    ts.SkipLine
                                        tempScale = Split(ts.ReadLine, ",")
                                    ts.SkipLine
                                        tempType = Val(ts.ReadLine)
                                    ts.SkipLine
                                        tempInteractiveFlags = Split(ts.ReadLine, ",")
                                    ts.SkipLine
                                        tempSolidFlags = Split(ts.ReadLine, ",")
                                    ts.SkipLine
                                        tempEnemyFlag = Val(ts.ReadLine)
                                    ts.SkipLine
                                    
                        If k = gameObjectLayer.layerBg2 Or k = gameObjectLayer.layerbg1 Or k = gameObjectLayer.layerfg1 Or k = gameObjectLayer.layerfg2 Then
                            map(mapIndex).addNewScenery Int(tempType), Int(k), Val(tempPos(0)), Val(tempPos(1)), Val(tempScale(0)), Val(tempScale(1))
                        ElseIf k = gameObjectLayer.layerPe Then
                            map(mapIndex).addNewEnemy Int(tempType), Int(k), Val(tempPos(0)), Val(tempPos(1)), Val(tempScale(0)), Val(tempScale(1))
                        ElseIf k = gameObjectLayer.layerIntv Then
                            map(mapIndex).addNewInteractive Int(tempType), Int(k), Val(tempPos(0)), Val(tempPos(1)), Val(tempScale(0)), Val(tempScale(1))
                        ElseIf k = gameObjectLayer.layerPf Then
                            map(mapIndex).addNewPlayfield Int(tempType), Int(k), Val(tempPos(0)), Val(tempPos(1)), Val(tempScale(0)), Val(tempScale(1))
                        End If
                    
                    If k = gameObjectLayer.layerPf Then
                        Set tempItem = map(mapIndex).getPlayfield(map(mapIndex).getCount(layerPf))
                    ElseIf k = gameObjectLayer.layerIntv Then
                         Set tempItem = map(mapIndex).getInteractive(map(mapIndex).getCount(layerIntv))
                    ElseIf k = gameObjectLayer.layerPe Then
                        Set tempItem = map(mapIndex).getEnemy(map(mapIndex).getCount(layerPe))
                     Else 'If tempType = gameObjectType.typeScenery Then
                        Set tempItem = map(mapIndex).getScenery(map(mapIndex).getCount(layerbg1))
                    End If
                    
                    
                        For j = 0 To UBound(tempInteractiveFlags)
                            tempItem.setFlagInteractiveValue j, CBool(tempInteractiveFlags(j))
                        Next j
                    
                        For j = 0 To UBound(tempSolidFlags)
                            tempItem.setFlagSolidValue j, CBool(tempSolidFlags(j))
                        Next j
                        
                            tempItem.setFlagEnemy (tempEnemyFlag)
                            
                    If k = gameObjectLayer.layerPf Then
                        map(mapIndex).setPlayfield (map(mapIndex).getCount(layerPf)), tempItem
                    ElseIf k = gameObjectLayer.layerIntv Then
                         map(mapIndex).setInteractive (map(mapIndex).getCount(layerIntv)), tempItem
                    ElseIf k = gameObjectLayer.layerPe Then
                        map(mapIndex).setEnemy (map(mapIndex).getCount(layerPe)), tempItem
                    Else
                     'ElseIf tempType = gameObjectlayer. Then
                        map(mapIndex).setScenery (map(mapIndex).getCount(layerbg1)), tempItem
                    End If
                    
                            'map(mapIndex).setItem map(mapIndex).count, tempItem
                    'End If
                Next i
                ts.SkipLine
            Else
                ts.SkipLine
                ts.SkipLine
            End If
          '  ts.SkipLine
        Next k
    ts.Close
    
    'frmEditor.refreshObjectListBox
End Sub
Function isScreenWidescreen() As Boolean
    Dim myReturn As Boolean
        myReturn = False
        
        Dim tempScreenMath As Single
            tempScreenMath = (Screen.width * Screen.TwipsPerPixelX) / (Screen.height * Screen.TwipsPerPixelY)
    
        If (tempScreenMath > 1.3) And (tempScreenMath < 1.4) Then
            myReturn = False
        Else
            myReturn = True
        End If
        
        isScreenWidescreen = myReturn
End Function
Sub loadGameMusic(Optional passedType As Single = 0)
    If loadedLevelMusic = False Then
        loopMusic = False
        frmGame.mediaPlayer.Command = "Close"
        frmGame.mediaPlayer.Command = "Stop"
            If passedType = 0 Then
                frmGame.mediaPlayer.FileName = App.path & "\midi\AzureLake.mid"
            End If
        loopMusic = False
        
        frmGame.mediaPlayer.Command = "Open"
        frmGame.mediaPlayer.Command = "Play"
    Else
        loopMusic = False
        frmGame.mediaPlayer.Command = "Stop"
        frmGame.mediaPlayer.Command = "Close"
            If passedType = 0 Then
                frmGame.mediaPlayer.FileName = App.path & "\midi\AzureLake.mid"
            End If
        loopMusic = False
        
        frmGame.mediaPlayer.Command = "Open"
        frmGame.mediaPlayer.Command = "Play"
    End If
    
        loadedLevelMusic = True
    
End Sub
Sub loadSimpleBackground(Optional passedType As Single = 0)
    If loadedLevelSimpleBG = False Then
         'If passedType = 0 Then
                frmSprites.picSimpleBG = LoadPicture(App.path & "\graphics\bg\azureSimpleTest.jpg")
          '  End If
    End If
        loadedLevelSimpleBG = True
End Sub
