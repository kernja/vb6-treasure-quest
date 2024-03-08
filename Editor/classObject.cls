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

Public objectType As gameObjectType
Public objectLayer As gameObjectLayer
Public id As Single

Private flagSolid(0 To 15) As Boolean
Private flagInteractive(0 To 31) As Boolean
Private flagEnemyType As gameEnemyType
Private objectWorldData As gameObjectWorldData


Private Sub Class_Initialize()
    enabled = False
    deleteMe = False
    
    objectType = gameObjectType.typeNothing
    objectLayer = gameObjectLayer.layerNothing
    flagEnemyType = -1
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


Sub renderMe(ByRef passedHandle As Long)
    
    If deleteMe = False Then
        If enabled = True Then
            If objectType = typeScenery Then
                cScenery.renderMe Me, passedHandle
            ElseIf objectType = typePlayfield Then
                cPlayfield.renderMe Me, passedHandle
            ElseIf objectType = typeEnemy Then
                cEnemy.renderMe Me, passedHandle
            ElseIf objectType = typeInteractive Then
                cInteractive.renderMe Me, passedHandle
            End If
        End If
    End If
    
End Sub
Public Sub copyScenery(ByRef passedObject As classObject)
    'If passedObject.objectType = typeScenery Then
    '    With objectWorldData
    '        .posX = passedObject.getPosX
    '        .posY = passedObject.getPosY
    '        .scaleX = passedObject.getScaleX
    '        .scaleY = passedObject.getScaleY
    '        .sizeHeight = passedObject.getSizeHeight
    '        .sizeWidth = passedObject.getSizeWidth
    '    End With
        
    '    objectType = typeScenery
    '        cScenery = New classScenery
    '        cScenery.Copy passedScenery
    'End If
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
Function getFlagSolidValue(Index As Single) As Boolean
    getFlagSolidValue = False
    
    If Index > -1 And Index <= UBound(flagSolid) Then
        getFlagSolidValue = flagSolid(Index)
    End If
End Function
Sub setFlagSolidValue(Index As Single, value As Boolean)
    If Index > -1 And Index <= UBound(flagSolid) Then
         flagSolid(Index) = value
    End If
End Sub
Sub setFlagInteractiveValue(Index As Single, value As Boolean)
    If Index > -1 And Index <= UBound(flagInteractive) Then
         flagInteractive(Index) = value
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
End Sub
Function getFlagInteractiveValue(Index As Single) As Boolean
    getFlagInteractiveValue = False
    
    If Index > -1 And Index <= UBound(flagInteractive) Then
        getFlagInteractiveValue = flagInteractive(Index)
    End If
End Function
Private Sub Class_Terminate()
    On Error Resume Next
        Set cPlayfield = Nothing
        Set cScenery = Nothing
        Set cInteractive = Nothing
        Set cEnemy = Nothing
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

