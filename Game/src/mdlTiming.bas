Attribute VB_Name = "mdlTiming"
Private Freq As Currency
Private StartTime As Currency
Private EndTime As Currency
Public TimeElapse As Double
Public TimeFactor As Double
Public Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Public Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long


Public Function FrameTiming(ByVal Action As Boolean)
    Select Case Action
        Case True
        QueryPerformanceFrequency Freq
        QueryPerformanceCounter StartTime
    Case False
        QueryPerformanceCounter EndTime
        TimeElapse = (EndTime - StartTime) / (Freq * 0.001)
    End Select
End Function

