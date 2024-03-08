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

Public Sub renderMe(ByRef passedObject As classObject, ByRef passedHandle As Long, ByVal renderOffsetX As Single, ByVal renderOffsetY As Single)
    
    If sceneryType <> -1 Then
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
                StretchBlt passedHandle, tempRenderX, tempRenderY, .getSizeWidth * mainRenderStretch(0), .getSizeHeight * mainRenderStretch(1), frmSprites.picScenerySprites(0).hdc, colIndex, 0, 64, 64, vbSrcAnd
                StretchBlt passedHandle, tempRenderX, tempRenderY, .getSizeWidth * mainRenderStretch(0), .getSizeHeight * mainRenderStretch(1), frmSprites.picScenerySprites(1).hdc, colIndex, 0, 64, 64, vbSrcPaint
            End If
        ElseIf currentGameMode = modeOnePlayer Or currentGameMode = modetwoplayercoop Then
            StretchBlt passedHandle, tempRenderX, tempRenderY, .getSizeWidth * mainRenderStretch(0), .getSizeHeight * mainRenderStretch(1), frmSprites.picScenerySprites(0).hdc, colIndex, 0, 64, 64, vbSrcAnd
            StretchBlt passedHandle, tempRenderX, tempRenderY, .getSizeWidth * mainRenderStretch(0), .getSizeHeight * mainRenderStretch(1), frmSprites.picScenerySprites(1).hdc, colIndex, 0, 64, 64, vbSrcPaint
        Else
              If (tempRenderY >= -128 * mainRenderStretch(1)) And tempRenderY <= frmGame.picRender(currentRenderIndex).height Then
                If (tempRenderX >= -128 * mainRenderStretch(1)) And tempRenderX <= frmGame.picRender(currentRenderIndex).width Then
                StretchBlt passedHandle, tempRenderX, tempRenderY, .getSizeWidth * mainRenderStretch(0), .getSizeHeight * mainRenderStretch(1), frmSprites.picScenerySprites(0).hdc, colIndex, 0, 64, 64, vbSrcAnd
                    StretchBlt passedHandle, tempRenderX, tempRenderY, .getSizeWidth * mainRenderStretch(0), .getSizeHeight * mainRenderStretch(1), frmSprites.picScenerySprites(1).hdc, colIndex, 0, 64, 64, vbSrcPaint
                End If
            End If
        End If
         
        End With
    End If
End Sub

Public Function getType() As Integer
    getType = sceneryType
End Function

