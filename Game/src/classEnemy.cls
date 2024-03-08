VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "classEnemy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private enemyType As gameEnemyType
Private animationFrame As Integer
Private colIndex As Integer

Public enemyLookLeft As Boolean
Public enemyFalling As Boolean
Public enemyGravityTime As Single
Public width As Single
Public height As Single

Public Sub initEnemy(ByVal passedEnemyType As gameEnemyType)
    enemyType = passedEnemyType
     colIndex = 64 * enemyType
        
        enemyLookLeft = True
        enemyFalling = False
        enemyGravityTime = 0
        
         
        width = 32
        height = 32
End Sub

Private Sub Class_Initialize()
    enemyType = gameEnemyType.typeNothing
    
         
        width = 32
        height = 32
End Sub
Public Sub renderMe(ByRef passedObject As classObject, ByRef passedHandle As Long, ByVal renderOffsetX As Single, ByVal renderOffsetY As Single)
Dim tempRenderX As Single
Dim tempRenderY As Single

   
    If enemyType <> -1 Then
        'moveEnemy
        With passedObject
        
        If currentGameMode = modeOnePlayer Or currentGameMode = modetwoplayercoop Then
            tempRenderX = .getPosX * mainRenderStretch(0)
            tempRenderY = .getPosY * mainRenderStretch(1)
        Else
            tempRenderX = (.getPosX - map(renderindex).renderOffsetX) * mainRenderStretch(0)
            tempRenderY = (.getPosY - map(renderindex).renderOffsetY) * mainRenderStretch(1)
        End If
       
        
        
        If currentGameMode = modeTwoPlayerVS Then
              If (tempRenderY >= -128 * mainRenderStretch(1)) And tempRenderY <= frmGame.picRender(currentRenderIndex).height Then
                StretchBlt passedHandle, tempRenderX, tempRenderY, .getSizeWidth * mainRenderStretch(0), .getSizeHeight * mainRenderStretch(1), frmSprites.picEnemySprites(0).hdc, colIndex, 0, 64, 64, vbSrcAnd
                StretchBlt passedHandle, tempRenderX, tempRenderY, .getSizeWidth * mainRenderStretch(0), .getSizeHeight * mainRenderStretch(1), frmSprites.picEnemySprites(1).hdc, colIndex, 0, 64, 64, vbSrcPaint
            End If
        ElseIf currentGameMode = modeOnePlayer Or currentGameMode = modetwoplayercoop Then
            StretchBlt passedHandle, tempRenderX, tempRenderY, .getSizeWidth * mainRenderStretch(0), .getSizeHeight * mainRenderStretch(1), frmSprites.picEnemySprites(0).hdc, colIndex, 0, 64, 64, vbSrcAnd
            StretchBlt passedHandle, tempRenderX, tempRenderY, .getSizeWidth * mainRenderStretch(0), .getSizeHeight * mainRenderStretch(1), frmSprites.picEnemySprites(1).hdc, colIndex, 0, 64, 64, vbSrcPaint
        Else
             If (tempRenderY >= -128 * mainRenderStretch(1)) And tempRenderY <= frmGame.picRender(currentRenderIndex).height Then
                If (tempRenderX >= -128 * mainRenderStretch(1)) And tempRenderX <= frmGame.picRender(currentRenderIndex).width Then
                    StretchBlt passedHandle, tempRenderX, tempRenderY, .getSizeWidth * mainRenderStretch(0), .getSizeHeight * mainRenderStretch(1), frmSprites.picEnemySprites(0).hdc, colIndex, 0, 64, 64, vbSrcAnd
                    StretchBlt passedHandle, tempRenderX, tempRenderY, .getSizeWidth * mainRenderStretch(0), .getSizeHeight * mainRenderStretch(1), frmSprites.picEnemySprites(1).hdc, colIndex, 0, 64, 64, vbSrcPaint
                End If
            End If
        End If
         
        End With
    End If
End Sub
Public Function getType() As Integer
    getType = enemyType
End Function
