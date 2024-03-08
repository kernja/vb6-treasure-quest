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
Private items() As classObject
Public count As Single
Public mapName As String
Public mapBackground As Integer
Public mapMusic As Integer

Private Sub Class_Initialize()
    count = -1
End Sub
Private Sub addNewItem()
    count = count + 1
        If count = 0 Then
            ReDim items(0)
            'Set items(0) = New classObject
        Else
            ReDim Preserve items(0 To count)
        End If
        
    Set items(count) = New classObject
End Sub
Sub addNewScenery(passedType As Integer, passedLayer As Integer, passedPosX As Single, passedPosY As Single, Optional passedScaleX As Single = 1, Optional passedScaleY As Single = 1)
    addNewItem
        With items(count)
            .id = count
            .setPosX passedPosX
            .setPosY passedPosY
            .setScaleX passedScaleX
            .setScaleY passedScaleY
            .newScenery passedType, passedLayer
        End With
End Sub
Sub addNewInteractive(passedType As Integer, passedLayer As Integer, passedPosX As Single, passedPosY As Single, Optional passedScaleX As Single = 1, Optional passedScaleY As Single = 1)
    addNewItem
        With items(count)
            .id = count
            .setPosX passedPosX
            .setPosY passedPosY
            .setScaleX passedScaleX
            .setScaleY passedScaleY
            .newInteractive passedType, passedLayer
        End With
End Sub
Sub addNewEnemy(passedType As Integer, passedLayer As Integer, passedPosX As Single, passedPosY As Single, Optional passedScaleX As Single = 1, Optional passedScaleY As Single = 1)
    addNewItem
        With items(count)
            .id = count
            .setPosX passedPosX
            .setPosY passedPosY
            .setScaleX passedScaleX
            .setScaleY passedScaleY
            .newEnemy passedType, passedLayer
        End With
End Sub
Sub addNewPlayfield(passedType As Integer, passedLayer As Integer, passedPosX As Single, passedPosY As Single, Optional passedScaleX As Single = 1, Optional passedScaleY As Single = 1)
    addNewItem
        With items(count)
            .id = count
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
    
       
       ' If currentSelectionMode > 0 Then
       For i = 0 To 7
       
             For j = 0 To count
                If items(j).objectLayer = i Then
                'If items(i).objectLayer = (currentSelectionMode - 1) Then
                    items(j).renderMe (passedHandle)
                'End If
                End If
            Next j
            
            'End If
        Next i
       ' End If

        
        'StretchBlt passedHandle, 0, 0, 64, 64, frmSprites.picPlaygroundSprites(0).hdc, 0, 0, 64, 64, vbSrcAnd
        'StretchBlt passedHandle, 0, 0, 64, 64, frmSprites.picPlaygroundSprites(1).hdc, 0, 0, 64, 64, vbSrcPaint
        
        'StretchBlt frmEditor.picRender.hdc, 0, 0, 64, 64, frmSprites.pcScenerySprites(0).GraphicCell(0), 0, 0, 64, 64, vbSrcAnd
        'StretchBlt frmEditor.picRender.hdc, 0, 0, 64, 64, frmSprites.pcScenerySprites(1).GraphicCell(0), 0, 0, 64, 64, vbSrcPaint
End Sub

Public Function getItem(Index As Integer) As classObject
    Set getItem = items(Index)
End Function

Public Sub setItem(Index As Integer, ByRef passedObject As classObject)
    Set items(Index) = passedObject
End Sub
