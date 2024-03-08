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
Public Sub initEnemy(ByVal passedEnemyType As gameEnemyType)
    enemyType = passedEnemyType
     colIndex = 64 * enemyType
End Sub

Private Sub Class_Initialize()
    enemyType = gameEnemyType.typeNothing
End Sub
Public Sub renderMe(ByRef passedObject As classObject, ByRef passedHandle As Long)
    
    If enemyType <> -1 Then
        With passedObject
        'MsgBox .getSizeHeight
          StretchBlt passedHandle, .getPosX * 0.5, .getPosY * 0.5, .getSizeWidth * 0.5, .getSizeHeight * 0.5, frmSprites.picEnemySprites(0).hdc, colIndex, 0, 64, 64, vbSrcAnd
          StretchBlt passedHandle, .getPosX * 0.5, .getPosY * 0.5, .getSizeWidth * 0.5, .getSizeHeight * 0.5, frmSprites.picEnemySprites(1).hdc, colIndex, 0, 64, 64, vbSrcPaint
        End With
    End If
End Sub

Public Function getType() As Integer
    getType = enemyType
End Function
