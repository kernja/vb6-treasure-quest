Attribute VB_Name = "mdlMain"
Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public progRender As Boolean
Public progExit As Boolean
Public currentSelectionMode As Integer

Public objFSO As FileSystemObject
Dim ts As TextStream

Public map As classMap

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
    
    progRender = True
    progExit = False
    
    Set objFSO = New FileSystemObject
    frmEditor.Show
    renderMe
End Sub

Sub quitProgram()
       
       Set objFSO = Nothing
       Set map = Nothing
       End
End Sub

Sub renderMe()
    Do While progExit = False
        frmEditor.picRender.Cls
        If progRender = True Then
            map.renderMe frmEditor.picRender.hdc
        End If
        
        DoEvents
        frmEditor.picRender.Refresh
        DoEvents
    Loop
    
    quitProgram
End Sub

Sub newMap()
    progRender = False
        Set map = New classMap
    progRender = True
End Sub

Sub saveMap(Optional path As String = "map.txt")

    Dim i As Single, j As Single, k As Single
    Dim count As Single
    Dim tempItem As classObject
    
    Set ts = objFSO.CreateTextFile(path, True, False)
    
    'write the map name
    ts.WriteLine ("//------------------------------------------------------------------")
    ts.WriteLine ("//TREASURE QUEST MAP FILE")
    ts.WriteLine ("//    ...do not edit manually. Or something dramatic will happen. :)")
    ts.WriteLine ("//------------------------------------------------------------------")
    ts.WriteLine ("//This is the map name.")
    ts.WriteLine ("//---------------------")
    ts.WriteLine (frmEditor.txtMapName.Text)
    ts.WriteLine ("")
    ts.WriteLine ("//---------------------------")
    ts.WriteLine ("//This is the map background.")
    ts.WriteLine ("//0 = 'Lake Overview'")
    ts.WriteLine ("//---------------------------")
        For i = 0 To 0
            If frmEditor.mnuBackgroundOption(i).Checked = True Then
                ts.WriteLine (i)
            End If
        Next i
        
    ts.WriteLine ("//---------------------------")
    ts.WriteLine ("//This is the map music")
    ts.WriteLine ("//0 = 'Azure Lake'")
    ts.WriteLine ("//---------------------------")
        For i = 0 To 0
            If frmEditor.mnuMusicOption(i).Checked = True Then
                ts.WriteLine (i)
            End If
        Next i
        
        
        ts.WriteLine ("")
        ts.WriteLine ("//-----------")
        ts.WriteLine ("//MAP OBJECTS")
    For k = 0 To 6
        ts.WriteLine ("//-----------")
        
        If k = 0 Then
            ts.WriteLine ("//bgTwo Items")
        ElseIf k = 1 Then
            ts.WriteLine ("//bgOne Items")
        ElseIf k = 2 Then
            ts.WriteLine ("//bgInteractive Items")
        ElseIf k = 3 Then
            ts.WriteLine ("//bgCharacterEnemy Items")
        ElseIf k = 4 Then
            ts.WriteLine ("//Playfield Items")
        ElseIf k = 5 Then
            ts.WriteLine ("//fgOne Items")
        Else
            ts.WriteLine ("//fgTwo Items")
        End If
        
        ts.WriteLine ("//-----------")
        ts.WriteLine ("//Count:")
        
        count = -1
            For i = 0 To map.count
                Set tempItem = map.getItem(Int(i))
                    If tempItem.deleteMe = False Then
                        If tempItem.objectLayer = k Then
                            count = count + 1
                        End If
                    End If
            Next i
            ts.WriteLine count
                If count > -1 Then
                    ts.WriteLine ("")
                    For i = 0 To map.count
                        
                        Set tempItem = map.getItem(Int(i))
                        
                        If tempItem.objectLayer = k Then
                        If tempItem.deleteMe = False Then
                            ts.WriteLine ("//------")
                            ts.WriteLine ("//item #" & i)
                            ts.WriteLine ("//------")
                        
                                
                                    ts.WriteLine ("//position:")
                                        ts.WriteLine (tempItem.getPosX & "," & tempItem.getPosY)
                                    ts.WriteLine ("//scale:")
                                        ts.WriteLine (tempItem.getScaleX & "," & tempItem.getScaleY)
                                    ts.WriteLine ("//item type:")
                                                If k = 0 Then
                                                    ts.WriteLine (tempItem.getSceneryType)
                                                ElseIf k = 1 Then
                                                    ts.WriteLine (tempItem.getSceneryType)
                                                ElseIf k = 2 Then
                                                    ts.WriteLine (tempItem.getInteractiveType)
                                                ElseIf k = 3 Then
                                                    ts.WriteLine (tempItem.getEnemyType)
                                                ElseIf k = 4 Then
                                                    ts.WriteLine (tempItem.getPlayfieldType)
                                                ElseIf k = 5 Then
                                                    ts.WriteLine (tempItem.getSceneryType)
                                                ElseIf k = 6 Then
                                                    ts.WriteLine (tempItem.getSceneryType)
                                                End If
                                    ts.WriteLine ("//interactive flags:")
                                        For j = 0 To tempItem.getFlagInteractiveCount
                                            ts.Write tempItem.getFlagInteractiveValue(j)
                                            
                                                If j < tempItem.getFlagInteractiveCount Then
                                                    ts.Write (",")
                                                Else
                                                    ts.WriteLine ("")
                                                End If
                                        Next j
                                    ts.WriteLine ("//solid flags:")
                                        For j = 0 To tempItem.getFlagSolidCount
                                            ts.Write tempItem.getFlagSolidValue(j)
                                            
                                                If j < tempItem.getFlagSolidCount Then
                                                    ts.Write (",")
                                                Else
                                                ts.WriteLine ("")
                                                End If
                                        Next j
                                    ts.WriteLine ("//enemy flag:")
                                        ts.WriteLine tempItem.getFlagEnemy
                                        ts.WriteLine ("")
                                End If
                                
                            
                        End If
                    Next i
                Else
                    ts.WriteLine ("")
                End If
              count = -1
            'ts.WriteLine ("")
        Next k
    ts.WriteLine ("")
    ts.Close
End Sub
Sub loadMap(Optional path As String = "map.txt")

    Dim i As Single, j As Single, k As Single, l As Single
    Dim count As Single
    Dim tempItem As classObject
    Dim tempType As Single
    Dim tempScale() As String
    Dim tempPos() As String
    Dim tempInteractiveFlags() As String
    Dim tempSolidFlags() As String
    Dim tempEnemy As Single
    
    Set map = New classMap
    Set ts = objFSO.OpenTextFile(path, ForReading, False, TristateFalse)
    
    'skip the six header lines
    For i = 0 To 5
        ts.SkipLine
    Next i
    'input the map name
        map.mapName = ts.ReadLine
            frmEditor.txtMapName = map.mapName
    'skip 5 more lines
    For i = 0 To 4
        ts.SkipLine
    Next i
    'input map background
        map.mapBackground = Val(ts.ReadLine)
            frmEditor.mnuBackgroundOption(map.mapBackground).Checked = True
    
    'skip 4 more lines
    For i = 0 To 4
        ts.SkipLine
    Next i
   'input map music
        map.mapMusic = Val(ts.ReadLine)
            frmEditor.mnuMusicOption(map.mapBackground).Checked = True
  
    'skip three more lines
        For i = 0 To 2
            ts.SkipLine
        Next i
    For k = 0 To 6
        'skip four more lines
        For i = 0 To 2
            'MsgBox ts.ReadLine 'ts.WriteLine ("//-----------")
            'MsgBox ts.ReadLine
            ts.SkipLine
        Next i
       ' MsgBox ts.ReadLine
           count = Val(ts.ReadLine)
                If count > -1 Then
                    For i = 0 To count
                        Set tempItem = New classObject
                          '  if tempitem.obje
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
                                       'MsgBox tempPos(0) & "," & tempPos(1)
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
                            map.addNewScenery Int(tempType), Int(k), Val(tempPos(0)), Val(tempPos(1)), Val(tempScale(0)), Val(tempScale(1))
                        ElseIf k = gameObjectLayer.layerPe Then
                            map.addNewEnemy Int(tempType), Int(k), Val(tempPos(0)), Val(tempPos(1)), Val(tempScale(0)), Val(tempScale(1))
                        ElseIf k = gameObjectLayer.layerIntv Then
                            map.addNewInteractive Int(tempType), Int(k), Val(tempPos(0)), Val(tempPos(1)), Val(tempScale(0)), Val(tempScale(1))
                        ElseIf k = gameObjectLayer.layerPf Then
                            map.addNewPlayfield Int(tempType), Int(k), Val(tempPos(0)), Val(tempPos(1)), Val(tempScale(0)), Val(tempScale(1))
                        End If
                        
                    Set tempItem = map.getItem(map.count)
                        For j = 0 To UBound(tempInteractiveFlags)
                            tempItem.setFlagInteractiveValue j, CBool(tempInteractiveFlags(j))
                        Next j
                    
                        For j = 0 To UBound(tempSolidFlags)
                            tempItem.setFlagSolidValue j, CBool(tempSolidFlags(j))
                        Next j
                        
                            tempItem.setFlagEnemy (tempEnemyFlag)
                            
                            map.setItem map.count, tempItem
                           
                Next i
                 
                ts.SkipLine
            Else
                ts.SkipLine
                ts.SkipLine
            End If
           ' ts.SkipLine
        Next k
    ts.Close
    
    'frmEditor.refreshObjectListBox
End Sub
