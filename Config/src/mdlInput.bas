Attribute VB_Name = "mdlInput"
Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Const KEY_TOGGLED As Integer = &H4
Const KEY_PRESSED As Integer = &H4000

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
