VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "classMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Private items() As classObject

Private sceneryAll() As classObject
Private sceneryInteractive() As classObject
Private sceneryEnemy() As classObject
Private sceneryPlayfield() As classObject

Private count(0 To 3) As Single
Public mapName As String
Public mapBackground As Integer
Public mapMusic As Integer
'Private gamePlayer(0 To 1) As Integer
Public renderOffsetX As Single
Public renderOffsetY As Single
Public newRenderOffsetX As Single
Public newRenderOffsetY As Single
Public renderOffsetXMax As Single
Public renderOffsetYMax As Single


Public Function getCount(layer As gameObjectLayer)
    If layer = 2 Then
        getCount = count(1)
    ElseIf layer = 3 Then
        getCount = count(2)
    ElseIf layer = 4 Then
        getCount = count(3)
    Else
        getCount = count(0)
    End If
End Function
Private Sub Class_Initialize()
    Dim i As Single
        For i = 0 To 3
            count(i) = -1
        Next i
End Sub
Private Sub addNewItem(passedValue As Single)
    Dim index As Single
        
        If passedValue <= 1 Or passedValue >= 5 Then index = 0
        If passedValue = 2 Then index = 1
        If passedValue = 3 Then index = 2
        If passedValue = 4 Then index = 3
        'End If
            count(index) = count(index) + 1
                
        If count(inde) = 0 Then
            If passedValue <= 1 Or passedValue >= 5 Then ReDim sceneryAll(0)
            If passedValue = 2 Then ReDim sceneryInteractive(0)
            If passedValue = 3 Then ReDim sceneryEnemy(0)
            If passedValue = 4 Then ReDim sceneryPlayfield(0)
            'Set items(0) = New classObject
        Else
            If passedValue <= 1 Or passedValue >= 5 Then ReDim Preserve sceneryAll(0 To count(index))
            If passedValue = 2 Then ReDim Preserve sceneryInteractive(0 To count(index))
            If passedValue = 3 Then ReDim Preserve sceneryEnemy(0 To count(index))
            If passedValue = 4 Then ReDim Preserve sceneryPlayfield(0 To count(index))
            'ReDim Preserve items(0 To count)
        End If
        
    If passedValue <= 1 Or passedValue >= 5 Then Set sceneryAll(count(index)) = New classObject
    If passedValue = 2 Then Set sceneryInteractive(count(index)) = New classObject
    If passedValue = 3 Then Set sceneryEnemy(count(index)) = New classObject
    If passedValue = 4 Then Set sceneryPlayfield(count(index)) = New classObject
    
    
End Sub
Sub addNewScenery(passedType As Integer, passedLayer As Integer, passedPosX As Single, passedPosY As Single, Optional passedScaleX As Single = 1, Optional passedScaleY As Single = 1)
    addNewItem (0)
        With sceneryAll(count(0))
            .id = count(0)
            .setPosX passedPosX
            .setPosY passedPosY
            .setScaleX passedScaleX
            .setScaleY passedScaleY
            .newScenery passedType, passedLayer
        End With
End Sub
Sub addNewInteractive(passedType As Integer, passedLayer As Integer, passedPosX As Single, passedPosY As Single, Optional passedScaleX As Single = 1, Optional passedScaleY As Single = 1)
    addNewItem (gameObjectLayer.layerIntv)
        With sceneryInteractive(count(1))
            .id = count(1)
            .setPosX passedPosX
            .setPosY passedPosY
            .setScaleX passedScaleX
            .setScaleY passedScaleY
            .newInteractive passedType, passedLayer
        End With
End Sub
Sub addNewEnemy(passedType As Integer, passedLayer As Integer, passedPosX As Single, passedPosY As Single, Optional passedScaleX As Single = 1, Optional passedScaleY As Single = 1)
    addNewItem (gameObjectLayer.layerPe)
        With sceneryEnemy(count(2))
            .id = count(2)
            .setPosX passedPosX
            .setPosY passedPosY
            .setScaleX passedScaleX
            .setScaleY passedScaleY
            .newEnemy passedType, passedLayer
        End With
End Sub
Sub addNewPlayfield(passedType As Integer, passedLayer As Integer, passedPosX As Single, passedPosY As Single, Optional passedScaleX As Single = 1, Optional passedScaleY As Single = 1)
    addNewItem (gameObjectLayer.layerPf)
        With sceneryPlayfield(count(3))
            .id = count(3)
            .setPosX passedPosX
            .setPosY passedPosY
            .setScaleX passedScaleX
            .setScaleY passedScaleY
            .newPlayfield passedType, passedLayer
        End With
End Sub
Public Sub renderMe(ByRef passedHandle As Long)
    Dim i As Single, j As Single
    Dim tempObject As classObject
    
       StretchBlt passedHandle, 0, 0, frmGame.picRender(currentRenderIndex).width, frmGame.picRender(currentRenderIndex).height, frmSprites.picSimpleBG.hdc, 0, 0, 1024, 768, vbSrcCopy

       ' If currentSelectionMode > 0 Then
       For i = 0 To 7
            
            If i = gameObjectLayer.layerPe Then
                   If Not (currentGameMode = modetwoplayercoop) Then
          '      Set mdlMain.players(currentLoadIndex) = New classPlayer
           '         With mdlMain.players(currentLoadIndex)
            '            .posX = objectWorldData.posX
             '           .posY = objectWorldData.posY
              '      End With
              
                mdlMain.players(currentRenderIndex).renderMe passedHandle, renderOffsetX, renderOffsetY
                    End If
            End If
          
            'End If
            If i <= 1 Or i >= 5 Then
                 For j = 0 To count(0)
                    If sceneryAll(j).objectLayer = i Then
                        sceneryAll(j).renderMe passedHandle, renderOffsetX, renderOffsetY
                    End If
                Next j
            ElseIf i = 2 Then
                For j = 0 To count(1)
                    If sceneryInteractive(j).enabled = True Then
                        sceneryInteractive(j).renderMe passedHandle, renderOffsetX, renderOffsetY
                    End If
                Next j
            ElseIf i = 3 Then
                For j = 0 To count(2)
                     If sceneryEnemy(j).enabled = True Then
                        sceneryEnemy(j).renderMe passedHandle, renderOffsetX, renderOffsetY
                    End If
                Next j
            ElseIf i = 4 Then
                For j = 0 To count(3)
                    If sceneryPlayfield(j).enabled = True Then
                        sceneryPlayfield(j).renderMe passedHandle, renderOffsetX, renderOffsetY
                    End If
                Next j
            End If

        Next i

        renderOffsetX = newRenderOffsetX
         renderOffsetY = newRenderOffsetY
End Sub

Public Function getScenery(index As Integer) As classObject
    Set getScenery = sceneryAll(index)
End Function

Public Sub setScenery(index As Integer, ByRef passedObject As classObject)
    Set sceneryAll(index) = passedObject
End Sub
Public Function getEnemy(index As Integer) As classObject
    Set getEnemy = sceneryEnemy(index)
End Function

Public Sub setEnemy(index As Integer, ByRef passedObject As classObject)
    Set sceneryEnemy(index) = passedObject
End Sub
Public Function getInteractive(index As Integer) As classObject
    Set getInteractive = sceneryInteractive(index)
End Function

Public Sub setInteractive(index As Integer, ByRef passedObject As classObject)
    Set sceneryInteractive(index) = passedObject
End Sub
Public Function getPlayfield(index As Integer) As classObject
    Set getPlayfield = sceneryPlayfield(index)
End Function

Public Sub setPlayfield(index As Integer, ByRef passedObject As classObject)
    Set sceneryPlayfield(index) = passedObject
End Sub
Public Function getCollideTop(ByRef player As classPlayer) As Single
    Dim myReturn As Single
        myReturn = -1
        
    Dim i As Single
    Dim tempItem As classObject
        For i = 0 To count(3)
            Set tempItem = sceneryPlayfield(i)
            
            If tempItem.enabled = True Then
                If (player.posX + player.width - 2 > tempItem.getPosX) And ((player.posX + 2) < tempItem.getPosX + tempItem.getSizeWidth) Then
                    If (player.posY + player.height > tempItem.getPosY) And ((player.posY + player.height) < tempItem.getPosY + tempItem.getSizeHeight) Then
                        myReturn = tempItem.getPosY
                    End If
                End If
            End If
        Next i
        
    getCollideTop = myReturn
End Function
Public Function getCollideBottom(ByRef player As classPlayer) As Single
    Dim myReturn As Single
        myReturn = -1
        
    Dim i As Single
    Dim tempItem As classObject
        For i = 0 To count(3)
            Set tempItem = sceneryPlayfield(i)
            
            If tempItem.enabled = True Then
                If (player.posX + player.width - 2 > tempItem.getPosX) And ((player.posX + 2) < tempItem.getPosX + tempItem.getSizeWidth) Then
                    If (player.posY < tempItem.getPosY + tempItem.getSizeHeight) And (player.posY > tempItem.getPosY) Then
                        myReturn = tempItem.getPosY + tempItem.getSizeHeight
                    End If
                End If
            End If
        Next i
        
    getCollideBottom = myReturn
End Function
Public Function getCollideLeft(ByRef player As classPlayer) As Single
    Dim myReturn As Single
        myReturn = -1
        
    Dim i As Single
    Dim tempItem As classObject
        For i = 0 To count(3)
            Set tempItem = sceneryPlayfield(i)
            
            If tempItem.enabled = True Then
                If (player.posX + 2 <= tempItem.getPosX + tempItem.getSizeWidth) And (player.posX + 2 >= tempItem.getPosX) Then
                    If (player.posY < tempItem.getPosY + tempItem.getSizeHeight - 1) And ((player.posY + player.height) > tempItem.getPosY + 1) Then
                        myReturn = tempItem.getPosX + tempItem.getSizeWidth
                    End If
                End If
            End If
        Next i
        
        
    getCollideLeft = myReturn
End Function
Public Function getCollideRight(ByRef player As classPlayer) As Single
    Dim myReturn As Single
        myReturn = -1
        
    Dim i As Single
    Dim tempItem As classObject
        For i = 0 To count(3)
            Set tempItem = sceneryPlayfield(i)
            
            If tempItem.enabled = True Then
                If (player.posX + player.width - 2 >= tempItem.getPosX) And ((player.posX + player.width - 2) <= tempItem.getPosX + tempItem.getSizeWidth) Then
                    If (player.posY < tempItem.getPosY + tempItem.getSizeHeight - 1) And ((player.posY + player.height) > tempItem.getPosY + 1) Then
                        myReturn = tempItem.getPosX - player.width
                    End If
                End If
            End If
        Next i
        
    getCollideRight = myReturn
End Function
Public Function getGameObjectCollideTop(ByRef GameObject As classObject) As Single
    Dim myReturn As Single
        myReturn = -1
        
    Dim i As Single
    Dim tempItem As classObject
        For i = 0 To count(3)
            Set tempItem = sceneryPlayfield(i)

            If tempItem.enabled = True Then
                If (GameObject.getPosX + GameObject.getSizeWidth - 2 > tempItem.getPosX) And ((GameObject.getPosX + 2) < tempItem.getPosX + tempItem.getSizeWidth) Then
                    If (GameObject.getPosY + GameObject.getSizeHeight > tempItem.getPosY) And ((GameObject.getPosY + GameObject.getSizeHeight) < tempItem.getPosY + tempItem.getSizeHeight) Then
                        myReturn = tempItem.getPosY
                    End If
                End If
            End If
        Next i
        
    getGameObjectCollideTop = myReturn
End Function
Public Function getGameObjectCollideBottom(ByRef GameObject As classObject) As Single
    Dim myReturn As Single
        myReturn = -1
        
    Dim i As Single
    Dim tempItem As classObject
        For i = 0 To count(3)
            Set tempItem = sceneryPlayfield(i)
            
            If tempItem.enabled = True Then
                If (GameObject.getPosX + GameObject.getSizeWidth - 2 > tempItem.getPosX) And ((GameObject.getPosX + 2) < tempItem.getPosX + tempItem.getSizeWidth) Then
                    If (GameObject.getPosY < tempItem.getPosY + tempItem.getSizeHeight) And (GameObject.getPosY > tempItem.getPosY) Then
                        myReturn = tempItem.getPosY + tempItem.getSizeHeight
                    End If
                End If
            End If
        Next i
        
    getGameObjectCollideBottom = myReturn
End Function
Public Function getGameObjectCollideLeft(ByRef GameObject As classObject) As Single
    Dim myReturn As Single
        myReturn = -1
        
    Dim i As Single
    Dim tempItem As classObject
        For i = 0 To count(3)
            Set tempItem = sceneryPlayfield(i)
            
            If tempItem.enabled = True Then
                If (GameObject.getPosX + 2 <= tempItem.getPosX + tempItem.getSizeWidth) And (GameObject.getPosX + 2 >= tempItem.getPosX) Then
                    If (GameObject.getPosY < tempItem.getPosY + tempItem.getSizeHeight - 1) And ((GameObject.getPosY + GameObject.getSizeHeight) > tempItem.getPosY + 1) Then
                        myReturn = tempItem.getPosX + tempItem.getSizeWidth
                    End If
                End If
            End If
        Next i
        
        
    getGameObjectCollideLeft = myReturn
End Function
Public Function getGameObjectCollideRight(ByRef GameObject As classObject) As Single
    Dim myReturn As Single
        myReturn = -1
        
    Dim i As Single
    Dim tempItem As classObject
        For i = 0 To count(3)
            Set tempItem = sceneryPlayfield(i)
            
            If tempItem.enabled = True Then
                If (GameObject.getPosX + GameObject.getSizeWidth - 2 >= tempItem.getPosX) And ((GameObject.getPosX + GameObject.getSizeWidth - 2) <= tempItem.getPosX + tempItem.getSizeWidth) Then
                    If (GameObject.getPosY < tempItem.getPosY + tempItem.getSizeHeight - 1) And ((GameObject.getPosY + GameObject.getSizeHeight) > tempItem.getPosY + 1) Then
                        myReturn = tempItem.getPosX - GameObject.getSizeWidth
                        
                    End If
                End If
            End If
        Next i
        
    getGameObjectCollideRight = myReturn
End Function
