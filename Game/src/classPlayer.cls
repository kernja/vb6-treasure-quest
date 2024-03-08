VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "classPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public posX As Single
Public posY As Single
Public width As Single
Public height As Single

Public playerLookLeft As Boolean
Public playerJumping As Boolean
Public playerFalling As Boolean
Public playerFlying As Boolean
Public playerGravityTime As Single
Public playerJumpTime As Single
Private animationFrame As Integer
Private colIndex As Integer

Public Sub initEnemy(ByVal passedEnemyType As gameEnemyType)

     width = 20
        height = 30
End Sub

Private Sub Class_Initialize()
    'enemyType = gameEnemyType.typeNothing
     width = 20
        height = 30
        
        gravityTime = 0
End Sub
Private Sub moveCharacter()
    Dim tempTopHeight As Single, tempBottomHeight As Single
    Dim tempLeftWidth As Single, tempRightWidth As Single
        
    
       tempLeftWidth = map(currentRenderIndex).getCollideLeft(Me)
        
        
            If tempLeftWidth <> -1 Then
                posX = tempLeftWidth ' + 32
            End If
               
        tempRightWidth = map(currentRenderIndex).getCollideRight(Me)
            If tempRightWidth <> -1 Then
                posX = tempRightWidth
            End If
            
        If playerFalling = True Then
            playerGravityTime = playerGravityTime + (1 * TimeElapse)
                If playerGravityTime > 10 Then playerGravityTime = 10
                    posY = posY + (playerGravityTime * 2 * TimeElapse)
        End If
                
        If playerJumping = True Then
            playerJumpTime = playerJumpTime - (1 * TimeElapse)
            posY = posY - (playerJumpTime * 2 * TimeElapse)
                If playerJumpTime < 0 Then
                    playerJumpTime = 0
                    playerJumping = False
                    playerFalling = True
                End If
        End If
        
        tempTopHeight = map(currentRenderIndex).getCollideTop(Me)
       
      
        
            If tempTopHeight = -1 Then
                If playerJumping = False And playerFlying = False Then
                    playerFalling = True
                End If
            Else
                If playerFalling = True Then
                    posY = tempTopHeight - height - 1
                        playerFalling = False
                        playerGravityTime = 0
                End If
            End If
            
             tempBottomHeight = map(currentRenderIndex).getCollideBottom(Me)
            If tempBottomHeight <> -1 Then
                If playerJumping = True Or playerFlying = True Then
                    posY = tempBottomHeight
                    
                        If playerJumping = True Then
                            playerJumping = False
                            playerFalling = True
                        End If
                    
                    jumptime = 0
                End If
            End If
                
     
End Sub
Private Sub getInput()
    If Not (currentGameMode = modetwoplayercoop) Then
       
       moveCharacter
       
    If posX < 0 Then posX = 0
    If posY < 0 Then
        posY = 0
            If playerJumping = True Then
                playerJumping = False
                playerFalling = True
            End If
        
        jumptime = 0
    End If
    
    If posX > 992 Then posX = 992
    If posY > 736 Then
        posY = 736
        playerGravityTime = 0
        playerFalling = False
    End If
    
            If gameInputMappings.inputUseKeyboard(currentRenderIndex) = True Then
                With gameInputMappings.inputMappingsKeyboard(currentRenderIndex)
                    If (GetKeyState(.keyDirectionDown) And KEY_DOWN) Then
                      '  posY = posY + 5 * TimeElapse
                    End If
                    If GetKeyState(.keyDirectionUp) And KEY_DOWN Then
                        'posY = posY - 5 * TimeElapse
                    End If
                    If GetKeyState(.keyDirectionLeft) And KEY_DOWN Then
                    posX = posX - 5 * TimeElapse
                    End If
                    If GetKeyState(.keyDirectionRight) And KEY_DOWN Then
                    posX = posX + 5 * TimeElapse
                    
                    End If
                    If GetKeyState(.keyFly) And KEY_DOWN Then
                         Beep
                    End If
                    If GetKeyState(.keyJump) And KEY_DOWN Then
                        'MsgBox playerGravityTime
                      If playerFalling = False And playerFlying = False And playerJumping = False Then
                        playerJumping = True
                        playerJumpTime = 10
                      End If
                    End If
                    If GetKeyState(.keyPause) And KEY_DOWN Then
                        Beep
                    End If
                    
                End With
            End If
    End If
    
    If posX < 0 Then posX = 0
    If posY < 0 Then posY = 0
    
    If posX > 992 Then posX = 992
    If posY > 736 Then posY = 736
End Sub
Public Sub renderMe(ByRef passedHandle As Long, ByVal renderOffsetX As Single, ByVal renderOffsetY As Single)
Dim tempRenderX As Single
Dim tempRenderY As Single

    getInput
    
    If enemyType <> -1 Then
        With passedObject
          If currentGameMode = modeOnePlayer Or currentGameMode = modetwoplayercoop Then
            tempRenderX = posX * mainRenderStretch(0)
            tempRenderY = posY * mainRenderStretch(1)
        Else
            tempRenderX = (posX - map(renderindex).renderOffsetX) * mainRenderStretch(0)
            tempRenderY = (posY - map(renderindex).renderOffsetY) * mainRenderStretch(1)
        End If
       
        
        If currentGameMode = modeTwoPlayerVS Then
            If (tempRenderY >= -128 * mainRenderStretch(1)) And tempRenderY <= frmGame.picRender(currentRenderIndex).height Then
                StretchBlt passedHandle, tempRenderX, tempRenderY, width * mainRenderStretch(0), height * mainRenderStretch(1), frmSprites.picPlayerSprite(0).hdc, colIndex, 0, 40, 60, vbSrcAnd
                StretchBlt passedHandle, tempRenderX, tempRenderY, width * mainRenderStretch(0), height * mainRenderStretch(1), frmSprites.picPlayerSprite(1).hdc, colIndex, 0, 40, 60, vbSrcPaint
            End If
        ElseIf currentGameMode = modeOnePlayer Or currentGameMode = modetwoplayercoop Then
            StretchBlt passedHandle, tempRenderX, tempRenderY, width * mainRenderStretch(0), height * mainRenderStretch(1), frmSprites.picPlayerSprite(0).hdc, colIndex, 0, 40, 60, vbSrcAnd
            StretchBlt passedHandle, tempRenderX, tempRenderY, width * mainRenderStretch(0), height * mainRenderStretch(1), frmSprites.picPlayerSprite(1).hdc, colIndex, 0, 40, 60, vbSrcPaint
        Else
            If (tempRenderY >= -128 * mainRenderStretch(1)) And tempRenderY <= frmGame.picRender(currentRenderIndex).height Then
                If (tempRenderX >= -128 * mainRenderStretch(1)) And tempRenderX <= frmGame.picRender(currentRenderIndex).width Then
                    StretchBlt passedHandle, tempRenderX, tempRenderY, width * mainRenderStretch(0), height * mainRenderStretch(1), frmSprites.picPlayerSprite(0).hdc, colIndex, 0, 40, 60, vbSrcAnd
                    StretchBlt passedHandle, tempRenderX, tempRenderY, width * mainRenderStretch(0), height * mainRenderStretch(1), frmSprites.picPlayerSprite(1).hdc, colIndex, 0, 40, 60, vbSrcPaint
                End If
            End If
        End If
         
        End With
    End If
    
    changeRenderOffset
End Sub
Sub changeRenderOffset()
    If Not (currentGameMode = modeOnePlayer) And Not (currentGameMode = modetwoplayercoop) Then
        With map(currentRenderIndex)
           If Not (currentGameMode = modeTwoPlayerVS) Then
            .newRenderOffsetX = posX * mainRenderStretch(0) - frmGame.picRender(currentRenderIndex).width * 0.5  ' * 0.5 ' - 32 * mainRenderStretch(0)
               If .newRenderOffsetX < 0 Then .newRenderOffsetX = 0
                If .newRenderOffsetX > frmGame.picRender(currentRenderIndex).width / mainRenderStretch(0) Then
                    .newRenderOffsetX = frmGame.picRender(currentRenderIndex).width / mainRenderStretch(0)
                End If
                
           End If
           
            .newRenderOffsetY = posY * mainRenderStretch(1) - frmGame.picRender(currentRenderIndex).height * 0.5
                If .newRenderOffsetY < 0 Then .newRenderOffsetY = 0
                If .newRenderOffsetY > (frmGame.picRender(currentRenderIndex).height) / mainRenderStretch(1) Then
                    .newRenderOffsetY = (frmGame.picRender(currentRenderIndex).height) / mainRenderStretch(1)
                End If
        End With
            
    End If
End Sub
Public Function getType() As Integer
    getType = enemyType
End Function
