VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "classScenery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private sceneryType As Integer
Private animationFrame As Integer
Private colIndex As Integer
Public Sub initScenery(ByVal passedSceneryType As Integer)
    sceneryType = passedSceneryType
        colIndex = 64 * sceneryType
End Sub

Private Sub Class_Initialize()
    sceneryType = 0
End Sub

Public Sub renderMe(ByRef passedObject As classObject, ByRef passedHandle As Long)
    
    If sceneryType <> -1 Then
        With passedObject
        'MsgBox .getSizeHeight
          StretchBlt passedHandle, .getPosX * 0.5, .getPosY * 0.5, .getSizeWidth * 0.5, .getSizeHeight * 0.5, frmSprites.picScenerySprites(0).hdc, colIndex, 0, 64, 64, vbSrcAnd
          StretchBlt passedHandle, .getPosX * 0.5, .getPosY * 0.5, .getSizeWidth * 0.5, .getSizeHeight * 0.5, frmSprites.picScenerySprites(1).hdc, colIndex, 0, 64, 64, vbSrcPaint
        End With
    End If
End Sub

Public Function getType() As Integer
    getType = sceneryType
End Function

