Attribute VB_Name = "mdlInput"
Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Const KEY_TOGGLED As Integer = &H4
Public Const KEY_DOWN As Long = -128

Public Type gameInputKeys
    keyJump As Long
    keyFly As Long
    keyPause As Long
    keyDirectionLeft As Long
    keyDirectionRight As Long
    keyDirectionUp As Long
    keyDirectionDown As Long
End Type

Public Type gameInputJoystick
    foo As Boolean
End Type

Public Type gameInputTotal
    inputUseKeyboard(0 To 3) As Boolean
    inputMappingsKeyboard(0 To 3) As gameInputKeys
    inputMappingsJoystick(0 To 3) As gameInputJoystick
End Type

Public gameInputMappings As gameInputTotal

Sub defineDefaultKeys()
    Dim i As Single
        For i = 0 To 3
            With gameInputMappings
                .inputUseKeyboard(i) = True
            End With
            
            With gameInputMappings.inputMappingsKeyboard(i)
                .keyFly = vbKeyShift
                .keyJump = vbKeyZ
                .keyPause = vbKeyReturn
                .keyDirectionDown = vbKeyDown
                .keyDirectionUp = vbKeyUp
                .keyDirectionLeft = vbKeyLeft
                .keyDirectionRight = vbKeyRight
            End With
        Next i
End Sub
