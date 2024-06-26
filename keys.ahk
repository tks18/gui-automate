#Include "%A_ScriptDir%\UI\index.ahk"

CapsLock & Space:: {
    ui := Interface()

    ui.gui.Show()
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