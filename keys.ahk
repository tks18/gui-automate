#Include "%A_ScriptDir%\UI\index.ahk"

CapsLock & Space:: {
    toolsGUI := Interface()
    addMainEditBox(toolsGUI)
    toolsGUI.ui.Show()
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