VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "classObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public enabled As Boolean
Public deleteMe As Boolean

Private cScenery As classScenery
Private cPlayfield As classPlayfield
Private cEnemy As classEnemy
Private cInteractive As classInteractive
Private cPlayer As classPlayer

Public objectType As gameObjectType
Public objectLayer As gameObjectLayer
Public id As Single

Private flagSolid(0 To 15) As Boolean
Private flagInteractive(0 To 19) As Boolean
Private flagEnemyType As gameEnemyType
Private objectWorldData As gameObjectWorldData


Private Sub Class_Initialize()
    enabled = False
    deleteMe = False
    
    objectType = gameObjectType.typeNothing
    objectLayer = gameObjectLayer.layerNothing
    
    'Exit Sub
    With objectWorldData
        .posX = 0
        .posY = 0
        .scaleX = 1
        .scaleY = 1
        .sizeHeight = 32
        .sizeWidth = 32
    End With
End Sub
Public Sub newEnemy(ByVal passedEnemyType As Integer, ByVal passedObjectLayer As gameObjectLayer)
    newObject gameObjectType.typeEnemy, passedObjectLayer
     Set cEnemy = New classEnemy
        cEnemy.initEnemy (passedEnemyType)
        objectLayer = passedObjectLayer
        enabled = True
        objectType = typeEnemy
    End Sub
Public Sub newInteractive(ByVal passedInteractiveType As Integer, ByVal passedObjectLayer As gameObjectLayer)
    newObject gameObjectType.typeInteractive, passedObjectLayer
     Set cInteractive = New classInteractive
        cInteractive.initInteractive (passedInteractiveType)
        objectLayer = passedObjectLayer
        enabled = True
        objectType = typeInteractive
    End Sub
Public Sub newScenery(ByVal passedSceneryType As Integer, ByVal passedObjectLayer As gameObjectLayer)
    newObject gameObjectType.typeScenery, passedObjectLayer
     Set cScenery = New classScenery
        cScenery.initScenery (passedSceneryType)
        objectLayer = passedObjectLayer
        enabled = True
        objectType = typeScenery
    End Sub
Public Sub newPlayfield(ByVal passedPlayfieldType As Integer, ByVal passedObjectLayer As gameObjectLayer)
    newObject gameObjectType.typePlayfield, passedObjectLayer
     Set cPlayfield = New classPlayfield
        cPlayfield.initPlayfield (passedPlayfieldType)
        objectLayer = passedObjectLayer
        objectType = typePlayfield
        enabled = True
    End Sub
Private Sub newObject(ByVal passedObjectType As gameObjectType, passedObjectLayer As gameObjectLayer)
    objectType = passedType
    objectLayer = passedObjectLayer
End Sub


Sub renderMe(ByRef passedHandle As Long, ByVal passedLimitX As Single, passedLimitY As Single)
    
    If deleteMe = False Then
        If enabled = True Then
            If objectType = typeScenery Then
                cScenery.renderMe Me, passedHandle, passedLimitX, passedLimitY
            ElseIf objectType = typePlayfield Then
                cPlayfield.renderMe Me, passedHandle, passedLimitX, passedLimitY
            ElseIf objectType = typeEnemy Then
                moveEnemy
                cEnemy.renderMe Me, passedHandle, passedLimitX, passedLimitY
            ElseIf objectType = typeInteractive Then
                cInteractive.renderMe Me, passedHandle, passedLimitX, passedLimitY
            End If
        End If
    End If
    
End Sub
Function getPosX() As Integer
    getPosX = objectWorldData.posX
End Function
Function getPosY() As Single
    getPosY = objectWorldData.posY
End Function
Sub setPosX(ByVal passedSingle As Single)
    objectWorldData.posX = passedSingle
End Sub
Sub setPosY(ByVal passedSingle As Single)
    objectWorldData.posY = passedSingle
End Sub
Function getScaleX() As Single
    getScaleX = objectWorldData.scaleX
End Function
Function getScaleY() As Single
    getScaleY = objectWorldData.scaleY
End Function
Sub setScaleX(ByVal passedSingle As Single)
    objectWorldData.scaleX = passedSingle
            objectWorldData.sizeWidth = 32 * passedSingle
End Sub
Sub setScaleY(ByVal passedSingle As Single)
    objectWorldData.scaleY = passedSingle
        objectWorldData.sizeHeight = 32 * passedSingle
End Sub
Function getSizeHeight() As Single
   getSizeHeight = objectWorldData.sizeHeight
End Function
Function getSizeWidth() As Single
    getSizeWidth = objectWorldData.sizeWidth
End Function
Function getFlagSolidCount() As Integer
    getFlagSolidCount = UBound(flagSolid)
End Function
Function getFlagSolidValue(index As Single) As Boolean
    getFlagSolidValue = False
    
    If index > -1 And index <= UBound(flagSolid) Then
        getFlagSolidValue = flagSolid(index)
    End If
End Function
Sub setFlagSolidValue(index As Single, value As Boolean)
    If index > -1 And index <= UBound(flagSolid) Then
         flagSolid(index) = value
    End If
End Sub
Sub setFlagInteractiveValue(index As Single, value As Boolean)
    If index > -1 And index <= UBound(flagInteractive) Then
         flagInteractive(index) = value
    End If
    
    If index = 17 Then
        If flagInteractive(17) = True Then
          If Not (currentGameMode = modetwoplayercoop) Then
                Set mdlMain.players(currentLoadIndex) = New classPlayer
                    With mdlMain.players(currentLoadIndex)
                        .posX = objectWorldData.posX
                        .posY = objectWorldData.posY
                    End With
                enabled = False
          End If
        End If
    End If
End Sub
Function getFlagInteractiveCount() As Integer
    getFlagInteractiveCount = UBound(flagInteractive)
End Function
Function getFlagEnemy() As gameEnemyType
    getFlagEnemy = flagEnemyType
End Function
Sub setFlagEnemy(value As gameEnemyType)
    flagEnemyType = value
    
    
    If value = 1 Then
    
    
    End If
End Sub
Function getFlagInteractiveValue(index As Single) As Boolean
    getFlagInteractiveValue = False
    
    If index > -1 And index <= UBound(flagInteractive) Then
        getFlagInteractiveValue = flagInteractive(index)
    End If
End Function
Private Sub Class_Terminate()
    On Error Resume Next
        Set cPlayfield = Nothing
        Set cScenery = Nothing
        Set cInteractive = Nothing
        Set cEnemy = Nothing
        Set cPlayer = Nothing
End Sub

Sub destroyObject()
    On Error Resume Next
        enabled = False
        deleteMe = True
        
        Set cScenery = Nothing
        Set cPlayfield = Nothing
        Set cInteractive = Nothing
        Set cEnemy = Nothing
End Sub

Public Function getSceneryType() As Integer
    getSceneryType = cScenery.getType
End Function
Public Function getEnemyType() As Integer
    getEnemyType = cEnemy.getType
End Function
Public Function getInteractiveType() As Integer
    getInteractiveType = cInteractive.getType
End Function
Public Function getPlayfieldType() As Integer
    getPlayfieldType = cPlayfield.getType
End Function
Public Sub moveEnemy()
    If flagEnemyType >= 0 Or flagenemytime <= 2 Then
       Dim tempTopHeight As Single, tempBottomHeight As Single
        Dim tempLeftWidth As Single, tempRightWidth As Single

        
            If cEnemy.enemyLookLeft = True Then
                If flagEnemyType = 0 Then objectWorldData.posX = objectWorldData.posX - 5 * TimeElapse
                If flagEnemyType = 1 Then objectWorldData.posX = objectWorldData.posX - 7.5 * TimeElapse
                If flagEnemyType = 2 Then objectWorldData.posX = objectWorldData.posX - 10 * TimeElapse
            Else
                If flagEnemyType = 0 Then objectWorldData.posX = objectWorldData.posX + 5 * TimeElapse
                If flagEnemyType = 1 Then objectWorldData.posX = objectWorldData.posX + 7.5 * TimeElapse
                If flagEnemyType = 2 Then objectWorldData.posX = objectWorldData.posX + 10 * TimeElapse
            End If
       tempLeftWidth = map(currentRenderIndex).getGameObjectCollideLeft(Me)
        
        
            If tempLeftWidth <> -1 Then
                objectWorldData.posX = tempLeftWidth ' + 32
                     cEnemy.enemyLookLeft = False
            End If
               
        tempRightWidth = map(currentRenderIndex).getGameObjectCollideRight(Me)
            If tempRightWidth <> -1 Then
                objectWorldData.posX = tempRightWidth
                    cEnemy.enemyLookLeft = True
            End If
            
            'If cEnemy.enemyFalling = False Then

            'End If
        If cEnemy.enemyFalling = True Then
             cEnemy.enemyGravityTime = cEnemy.enemyGravityTime + (1 * TimeElapse)
                If cEnemy.enemyGravityTime > 10 Then cEnemy.enemyGravityTime = 10
                    objectWorldData.posY = objectWorldData.posY + (cEnemy.enemyGravityTime * 2 * TimeElapse)
        End If
        
        tempTopHeight = map(currentRenderIndex).getGameObjectCollideTop(Me)
       
      
            If tempTopHeight = -1 Then
                     cEnemy.enemyFalling = True
            Else
                If cEnemy.enemyFalling = True Then
                   objectWorldData.posY = tempTopHeight - getSizeHeight - 1
                         cEnemy.enemyFalling = False
                         cEnemy.enemyGravityTime = 0
                End If
            End If
            
             tempBottomHeight = map(currentRenderIndex).getGameObjectCollideBottom(Me)
            If tempBottomHeight <> -1 Then
                    'keep this code?
            End If
            
    
     If objectWorldData.posX < 0 Then
        objectWorldData.posX = 0
        cEnemy.enemyLookLeft = False
    End If
    If objectWorldData.posY < 0 Then
        objectWorldData.posY = 0
    End If
    
    If objectWorldData.posX > 992 Then
        objectWorldData.posX = 992
        cEnemy.enemyLookLeft = True
    End If
    If objectWorldData.posY > 736 Then
        objectWorldData.posY = 736
        playerGravityTime = 0
        playerFalling = False
    End If
    
    
    End If
End Sub
