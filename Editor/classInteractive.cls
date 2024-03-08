VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "classInteractive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private interactiveType As gameInteractiveType
Private animationFrame As Integer
Private colIndex As Integer
Public Sub initInteractive(ByVal passedInteractiveType As gameInteractiveType)
    interactiveType = passedInteractiveType
     colIndex = 64 * interactiveType
End Sub

Private Sub Class_Initialize()
    interactiveType = gameInteractiveType.typeNothing
End Sub
Public Sub renderMe(ByRef passedObject As classObject, ByRef passedHandle As Long)
    
    If interactiveType <> -1 Then
        With passedObject
        'MsgBox .getSizeHeight
          StretchBlt passedHandle, .getPosX * 0.5, .getPosY * 0.5, .getSizeWidth * 0.5, .getSizeHeight * 0.5, frmSprites.picInteractiveSprites(0).hdc, colIndex, 0, 64, 64, vbSrcAnd
          StretchBlt passedHandle, .getPosX * 0.5, .getPosY * 0.5, .getSizeWidth * 0.5, .getSizeHeight * 0.5, frmSprites.picInteractiveSprites(1).hdc, colIndex, 0, 64, 64, vbSrcPaint
        End With
    End If
End Sub

Public Function getType() As Integer
    getType = interactiveType
End Function
