Private Declare PtrSafe Function SystemParametersInfo Lib "user32" Alias _
    "SystemParametersInfoA" (ByVal uAction As Integer, ByVal uParam As Integer, _
    ByVal lpvParam As Boolean, ByVal fuWinIni As Integer) As Integer

Sub UpdateMouse()
    Const AWT = &H1001
    Const UI = &H1
    Const WU = &H2
    
    Dim state As Boolean
    
    state = True
    
    Result = SystemParametersInfo(AWT, 0, state, WU)
End Sub
