#Requires AutoHotkey v2.0
SendMode("Input")
SetWorkingDir(A_ScriptDir)
#SingleInstance

#Include "%A_ScriptDir%\GUI\index.ahk"

SetCapsLockState("AlwaysOff")

CapsLock & Space:: {
    toolsGUI := BaseUI()
    addMainEditBox(toolsGUI)
    toolsGUI.BaseGUI.Show()
    return
}

!CapsLock:: {
    capsstate := GetKeyState("CapsLock", "T") ? "D" : "U"
    if (capsstate = "U")
        SetCapsLockState("AlwaysOn")
    else
        SetCapsLockState("AlwaysOff")
    return
}