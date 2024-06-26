#Include "%A_ScriptDir%\UI\index.ahk"

CapsLock & Space:: {
    toolsGUI := BaseUI()
    addMainEditBox(toolsGUI)
    toolsGUI.BaseGUI.Show()
    return
}

!CapsLock:: {
    capsState := GetKeyState("CapsLock", "T") ? "D" : "U"
    if (capsState = "U")
        SetCapsLockState("AlwaysOn")
    else
        SetCapsLockState("AlwaysOff")
    return
}